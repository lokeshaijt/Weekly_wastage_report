"""
Transaction Report parser (.xlsb or .xlsx export, header row located by content).

Per (item, date) it accumulates six quantities:
  0 consumption  PROCESS  / STOCKOUT
  1 floor        WASTAGE  / STOCKOUT, From Warehouse not a PM store
  2 store        WASTAGE  / STOCKOUT, From Warehouse in WH-PM-MDKCBE / WH-PM-MDKCBE2
  3 writeoff     WRITEOFF / STOCKOUT
  4 by-product   PROCESS  / STOCKOUT flagged "Is By Product Consumption" = Y   (used for blends)
  5 reclaimed    PROCESS  / STOCKIN  on work centre RECLAIM                     (used for blends)

Only the STOCKOUT leg is used for wastage/writeoff: each of those transactions also has a
paired STOCKIN leg into the wastage/writeoff warehouse, counting it would double the figures.
"""
import collections
from datetime import date, datetime, timedelta

from .utils import norm_disp

STORE_WAREHOUSES = {"WH-PM-MDKCBE", "WH-PM-MDKCBE2"}
CONS, FLOOR, STORE, WRITEOFF, BYPROD, RECLAIM = range(6)
REQUIRED = ("Transaction Type", "Item_Name", "Movement Type", "Doc_Date", "Qty")


def _rows_xlsb(path):
    import pyxlsb
    with pyxlsb.open_workbook(path) as wb:
        with wb.get_sheet(wb.sheets[0]) as sh:
            it = sh.rows()
            head = None
            for row in it:
                vals = [c.v for c in row]
                if all(h in vals for h in REQUIRED):
                    head = {str(x).strip(): i for i, x in enumerate(vals) if x}
                    break
            if head is None:
                raise ValueError(f"{path}: Transaction Report header row not found")
            for row in it:
                vals = [c.v for c in row]
                g = {k: (vals[i] if i < len(vals) else None) for k, i in head.items()}
                d = g.get("Doc_Date")
                g["Doc_Date"] = date(1899, 12, 30) + timedelta(days=int(d)) if isinstance(d, (int, float)) else None
                yield g


def _rows_xlsx(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    it = ws.iter_rows(values_only=True)
    head = None
    for r in it:
        if r and all(h in r for h in REQUIRED):
            head = {str(x).strip(): i for i, x in enumerate(r) if x}
            break
    if head is None:
        raise ValueError(f"{path}: Transaction Report header row not found")
    for r in it:
        if not r:
            continue
        g = {k: (r[i] if i < len(r) else None) for k, i in head.items()}
        d = g.get("Doc_Date")
        g["Doc_Date"] = d.date() if isinstance(d, datetime) else d
        yield g


def read_rows(path):
    return _rows_xlsb(path) if str(path).lower().endswith(".xlsb") else _rows_xlsx(path)


def parse(paths):
    """Returns (daily, meta): daily[(item, date)] = [6 floats]; meta[item] = (uom, item_type)."""
    daily = collections.defaultdict(lambda: [0.0] * 6)
    meta = {}
    for p in paths:
        for g in read_rows(p):
            tt = str(g.get("Transaction Type") or "").upper()
            mv = str(g.get("Movement Type") or "").upper()
            nm = norm_disp(g.get("Item_Name"))
            d = g.get("Doc_Date")
            if not nm or not isinstance(d, date) or tt not in ("PROCESS", "WASTAGE", "WRITEOFF"):
                continue
            try:
                q = float(g.get("Qty"))
            except (TypeError, ValueError):
                continue
            meta.setdefault(nm, (g.get("UOM_Name") or "", str(g.get("Item_Type_Name") or "").upper()))
            a = daily[(nm, d)]
            if mv == "STOCKOUT":
                if tt == "PROCESS":
                    a[CONS] += q
                    if str(g.get("Is By Product Consumption") or "").upper() == "Y":
                        a[BYPROD] += q
                elif tt == "WRITEOFF":
                    a[WRITEOFF] += q
                elif norm_disp(g.get("From Warehouse")).upper() in STORE_WAREHOUSES:
                    a[STORE] += q
                else:
                    a[FLOOR] += q
            elif mv == "STOCKIN" and tt == "PROCESS" and str(g.get("Workcenter_Name") or "").upper() == "RECLAIM":
                a[RECLAIM] += q
    return dict(daily), meta
