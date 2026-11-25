"""
Excel builder: Wastage Summary, Black Tea (BLENDB) and Speciality Tea (BLENDS), zone-wise.

Column blocks (left to right):
  each month of the summary period except the last (collapsed),
  the summary period total, the last month of the period, the current month to date, the last week.
Every % = wastage / (wastage + consumption); 0% if consumption <= 0 or wastage <= 0 (never > 100%),
on item rows and on Total / Zone Total / Grand Total rows alike. Red when >= 2%.
"""
import collections
import copy
import re

import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

from .transactions import BYPROD, CONS, FLOOR, RECLAIM, STORE, WRITEOFF

WASTE_COLS = ["Consumption", "Floor Wastage", "Floor Wastage %", "Store Wastage", "Store Wastage %",
              "Writeoff", "Writeoff %", "Total Wastage", "Total Wastage %"]
WASTE_SUM = ["Consumed Qty", "Floor Wastage %", "Store Wastage %", "Writeoff %", "Total Wastage", "Total Wastage %"]
BLEND_COLS = ["Consumption", "Floor Wastage", "Floor Wastage %", "Store Wastage", "Store Wastage %",
              "Wastage before Reclaim", "Wastage before Reclaim %", "By Prod Cons Qty", "Reclaimed Qty",
              "By Prod after Reclaim", "By Prod after Reclaim %", "Wastage after Reclaim", "Wastage after Reclaim %"]
BLEND_SUM = ["Consumed Qty"] + BLEND_COLS[1:]

# ---- look (taken from the agreed report format)
FONT = "Calibri"
TITLE_FONT = Font(name=FONT, size=11, bold=True, color="FFFF0000")
HDR_FONT = Font(name=FONT, size=9, bold=True, color="FFFFFFFF")
TXT_FONT = Font(name=FONT, size=9)
ITEM_HDR = "FF1F4E78"
MONTH_R1 = ["FF953735", "FFE46C0A"]                  # alternating block title colours
WASTE_R2 = ["FF403152", "FF254061"]                  # alternating column-header colours
BLEND_R2 = ["FF17375E", "FF4F6228"]
WASTE_SUM_FILL = ("FF4F6228", "FF4F6228")
BLEND_SUM_FILL = ("FF7E0C53", "FF20624E")
TOTAL_FILL = PatternFill("solid", fgColor="FFDCE6F1")
ZONE_FILL = PatternFill("solid", fgColor="FFFFD966")
GRAND_FILL = PatternFill("solid", fgColor="FFF4B183")
THIN = Side(style="thin", color="FFBFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
QTY_FMT = '_ * #,##0_ ;_ * \\-#,##0_ ;_ * "-"??_ ;_ @_ '
PCT_FMT = "0.0%"
RED = CellIsRule(operator="greaterThanOrEqual", formula=["0.02"], font=Font(color="FF9C0006"),
                 fill=PatternFill("solid", fgColor="FFFFC7CE", bgColor="FFFFC7CE"))
# columns left visible inside the visible blocks (the rest are grouped & collapsed)
VISIBLE = {"waste": {"Consumption", "Consumed Qty", "Total Wastage", "Total Wastage %"},
           "blend": {"Consumption", "Consumed Qty", "Wastage before Reclaim %", "Wastage after Reclaim %"}}
WIDTH = {"Consumption": 9.78, "Consumed Qty": 11.0, "Total Wastage": 7.78}


def pct(m, c):
    return f"=IF(OR({c}<=0,{m}<=0),0,{m}/({m}+{c}))"


def q(v):
    v = round(v, 4)
    return v if abs(v) > 1e-9 else None


class Period:
    """One column block: label + a bucket index into item data (None for the summary block)."""
    def __init__(self, key, label, bucket=None):
        self.key, self.label, self.bucket = key, label, bucket


def build_sheet(wb, name, title, kind, items, periods, zone_desc, collapse_rows=True):
    """
    items: list of (zone, category, item, uom, item_type, buckets) where buckets[i] = [6 floats]
    periods: ordered list of Period; 'month' periods with collapsed=True are grouped/hidden.
    """
    ws = wb.create_sheet(name)
    mh = WASTE_COLS if kind == "waste" else BLEND_COLS
    sh = WASTE_SUM if kind == "waste" else BLEND_SUM
    r2fill = WASTE_R2 if kind == "waste" else BLEND_R2
    sumfill = WASTE_SUM_FILL if kind == "waste" else BLEND_SUM_FILL
    # ---- header
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = Alignment(wrap_text=True, vertical="center")
    for j, (h, w) in enumerate(zip(["Zone", "Item Category", "Item Name", "UOM", "Item Type"],
                                   [9.0, 9.78, 21.66, 4.66, 9.89]), 1):
        c = ws.cell(2, j, h)
        c.font, c.fill, c.border = HDR_FONT, PatternFill("solid", fgColor=ITEM_HDR), BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[L(j)].width = w
    pos, col, alt = {}, 6, 0
    for p in periods:
        cols = sh if p.key == "sum" else mh
        pos[p.key] = col
        f1, f2 = (sumfill if p.key == "sum" else (MONTH_R1[alt % 2], r2fill[alt % 2]))
        if p.key != "sum":
            alt += 1
        ws.cell(1, col, p.label)
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + len(cols) - 1)
        for j, h in enumerate(cols):
            c1, c2 = ws.cell(1, col + j), ws.cell(2, col + j, h)
            for c, f in ((c1, f1), (c2, f2)):
                c.font, c.fill, c.border = HDR_FONT, PatternFill("solid", fgColor=f), BORDER
            c1.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c2.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True, text_rotation=90)
            cd = ws.column_dimensions[L(col + j)]
            cd.width = WIDTH.get(h, 7.33 if not h.endswith("%") else 5.5)
            if getattr(p, "collapsed", False):
                cd.outline_level = 1 if h in VISIBLE[kind] or h in ("Wastage before Reclaim", "Wastage after Reclaim") else 2
                cd.hidden = True
            elif h not in VISIBLE[kind]:
                cd.outline_level, cd.hidden = 1, True
        col += len(cols)
    last = col - 1
    ws.row_dimensions[1].height = 28.2
    ws.row_dimensions[2].height = 73.2 if kind == "waste" else 61.2
    pct_cols = [c for c in range(6, last + 1) if str(ws.cell(2, c).value).endswith("%")]
    month_keys = [p.key for p in periods if re.fullmatch(r"m\d+", p.key)]      # summary = the monthly blocks only

    def bcols(key):
        cols = sh if key == "sum" else mh
        return {h: pos[key] + j for j, h in enumerate(cols)}

    def pct_row(r):
        for p in periods:
            bc = bcols(p.key)
            for h, c in bc.items():
                if not h.endswith("%"):
                    continue
                base = h[:-1].strip()
                C = f"{L(pos[p.key])}{r}"
                if base in bc:
                    m = f"{L(bc[base])}{r}"
                else:                                                      # summary % without a qty column
                    m = "(" + "+".join(f"{L(bcols(k)[base])}{r}" for k in month_keys) + ")"
                ws.cell(r, c).value = pct(m, C)

    def write_item(r, data):
        parts = collections.defaultdict(list)
        for p in periods:
            if p.key == "sum":
                continue
            bc = bcols(p.key)
            A = lambda h: f"{L(bc[h])}{r}"
            v = data[p.bucket]
            ws[A("Consumption")], ws[A("Floor Wastage")], ws[A("Store Wastage")] = q(v[CONS]), q(v[FLOOR]), q(v[STORE])
            if kind == "waste":
                ws[A("Writeoff")] = q(v[WRITEOFF])
                ws[A("Total Wastage")] = f"={A('Floor Wastage')}+{A('Store Wastage')}+{A('Writeoff')}"
            else:
                ws[A("By Prod Cons Qty")], ws[A("Reclaimed Qty")] = q(v[BYPROD]), q(v[RECLAIM])
                ws[A("Wastage before Reclaim")] = f"={A('Floor Wastage')}+{A('Store Wastage')}+{A('By Prod Cons Qty')}"
                ws[A("By Prod after Reclaim")] = f"={A('By Prod Cons Qty')}-{A('Reclaimed Qty')}"
                ws[A("Wastage after Reclaim")] = f"={A('Floor Wastage')}+{A('Store Wastage')}+{A('By Prod after Reclaim')}"
            if p.key in month_keys:
                for h in mh:
                    if not h.endswith("%"):
                        parts[h].append(A(h))
        sc = bcols("sum")
        for h in sh:
            if not h.endswith("%"):
                ws.cell(r, sc[h]).value = "=" + "+".join(parts["Consumption" if h == "Consumed Qty" else h])
        pct_row(r)

    def style(r, fill=None, bold=False):
        for c in range(1, last + 1):
            x = ws.cell(r, c)
            x.font = Font(name=FONT, size=9, bold=bold)
            x.border = BORDER
            if c >= 6:
                x.number_format = PCT_FMT if c in pct_cols else QTY_FMT
            if fill:
                x.fill = copy.copy(fill)

    def total_row(r, a, b, label, fill, f_, level, hidden):
        for j, v in enumerate([a, b, label], 1):
            ws.cell(r, j, v)
        for c in range(6, last + 1):
            if c not in pct_cols:
                ws.cell(r, c).value = f_(L(c))
        pct_row(r)
        style(r, fill, True)
        ws.row_dimensions[r].outline_level = level
        ws.row_dimensions[r].hidden = hidden

    tree = collections.OrderedDict((z, collections.defaultdict(list)) for z in sorted({i[0] for i in items}))
    for it in items:
        tree[it[0]][it[1]].append(it)
    r, zone_rows, counts = 3, [], collections.Counter()
    for z, cats in tree.items():
        cat_rows = []
        for cat in sorted(cats):
            first = r
            for it in sorted(cats[cat], key=lambda t: t[2]):
                for j, v in enumerate([f"Zone {z}", cat, it[2], it[3], it[4]], 1):
                    ws.cell(r, j, v)
                write_item(r, it[5])
                style(r)
                ws.row_dimensions[r].outline_level = 2
                ws.row_dimensions[r].hidden = collapse_rows
                r += 1
                counts[z] += 1
            total_row(r, f"Zone {z}", cat, "Total", TOTAL_FILL, lambda C, a=first, b=r - 1: f"=SUM({C}{a}:{C}{b})", 1, collapse_rows)
            cat_rows.append(r)
            r += 1
        total_row(r, f"Zone {z}", zone_desc.get(z, ""), "Zone Total", ZONE_FILL,
                  lambda C, cr=tuple(cat_rows): "=" + "+".join(f"{C}{x}" for x in cr), 0, False)
        zone_rows.append(r)
        r += 1
    total_row(r, "All Zones", None, "Grand Total", GRAND_FILL,
              lambda C: "=" + "+".join(f"{C}{x}" for x in zone_rows), 0, False)
    ws.conditional_formatting.add(" ".join(f"{L(c)}3:{L(c)}{r}" for c in pct_cols), RED)
    ws.freeze_panes = "F3"
    ws.auto_filter.ref = f"A2:{L(last)}{r}"
    ws.sheet_properties.outlinePr.summaryBelow = True
    ws.sheet_properties.outlinePr.summaryRight = True
    return ws, counts


def add_unzoned_sheet(wb, rows):
    ws = wb.create_sheet("Items without Zone")
    ws["A1"] = "Items with wastage/consumption but no zone - enter the Zone (1-9 or zone name) in column A and pass this file with --item-zones"
    ws["A1"].font = TITLE_FONT
    hdr = ["Zone (fill in)", "Item Category", "Item Name", "UOM", "Item Type", "Consumption (summary period)", "Wastage (summary period)"]
    for j, (h, w) in enumerate(zip(hdr, [10, 12, 45, 6, 12, 14, 12]), 1):
        c = ws.cell(2, j, h)
        c.font, c.fill, c.border = HDR_FONT, PatternFill("solid", fgColor=ITEM_HDR), BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[L(j)].width = w
    for i, row in enumerate(rows, 3):
        for j, v in enumerate(row, 1):
            c = ws.cell(i, j, v)
            c.font, c.border = TXT_FONT, BORDER
            if j >= 6:
                c.number_format = "#,##0"
        ws.cell(i, 1).fill = PatternFill("solid", fgColor="FFFFFFCC")
    ws.freeze_panes = "A3"
    return ws
