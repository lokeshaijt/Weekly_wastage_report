"""
Command line entry point.

  python -m wastage_report \
      --history  Transaction_Report_Oct2025-Oct2026.xlsb \
      --current  TR_02Nov-07Nov.xlsx TR_WEEK_46.xlsx TR_WEEK_47.xlsx \
      --zone-master Item_Master_Zone_wise.xlsx \
      --item-zones  Zone_Master_Items_9Zones.xlsx \
      --bom Exploded_BOM.xlsx Exploded_BOM_Supplemental.xlsx Additional_bom.xlsx \
      --categories  previous_report.xlsx \
      --out Wastage_Summary.xlsx
"""
import argparse
import collections
import datetime as dt
import sys

import openpyxl

from . import __version__
from .masters import Masters
from .report import Period, add_unzoned_sheet, build_sheet
from .transactions import CONS, FLOOR, STORE, WRITEOFF, parse
from .utils import week_bounds, week_number_for

BLEND_SHEETS = {"BLENDB": "Black Tea", "BLENDS": "Speciality Tea"}


def month_range(a, b):
    y, m = a
    out = []
    while (y, m) <= b:
        out.append((y, m))
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return out


def ym(s):
    y, m = s.split("-")
    return int(y), int(m)


def main(argv=None):
    ap = argparse.ArgumentParser(prog="wastage_report", description="Zone-wise weekly wastage report (Wastage Summary, Black Tea, Speciality Tea)")
    ap.add_argument("--history", nargs="+", required=True, help="Transaction Report(s) covering the summary months (.xlsb/.xlsx)")
    ap.add_argument("--current", nargs="+", required=True, help="Transaction Reports for the current month, one per week; the latest week becomes the week block")
    ap.add_argument("--zone-master", required=True, help="FG zone master (sheets 'Active FG-12 Months' and 'Master')")
    ap.add_argument("--item-zones", nargs="*", default=[], help="Item-level zone files (Item Name | Zone); take priority over the BOM mapping")
    ap.add_argument("--bom", nargs="+", required=True, help="Exploded BOM / BOM-Routing-Grid files, oldest first (later files override an FG's BOM)")
    ap.add_argument("--categories", nargs="*", default=[], help="Item category sources: a previous report or an Item Name | Item Category list")
    ap.add_argument("--summary-start", help="First summary month YYYY-MM (default: first month in --history)")
    ap.add_argument("--summary-end", help="Last summary month YYYY-MM (default: last month in --history)")
    ap.add_argument("--out", required=True, help="Output .xlsx")
    ap.add_argument("--unzoned-sheet", action="store_true", help="Add an 'Items without Zone' sheet to fill in")
    ap.add_argument("--expand-rows", action="store_true", help="Open with item rows expanded (default: collapsed to Zone Totals)")
    a = ap.parse_args(argv)

    print(f"wastage_report {__version__}")
    hist, hmeta = parse(a.history)
    cur, cmeta = parse(a.current)
    cur_dates = {d for _, d in cur}
    if not cur_dates:
        sys.exit("no rows in --current files")
    hist = {k: v for k, v in hist.items() if k[1] not in cur_dates}          # current files win on overlapping dates
    meta = {**hmeta, **cmeta}
    hmonths = sorted({(d.year, d.month) for _, d in hist})
    start = ym(a.summary_start) if a.summary_start else hmonths[0]
    end = ym(a.summary_end) if a.summary_end else hmonths[-1]
    months = month_range(start, end)
    last_day = max(cur_dates)
    wk = week_number_for(last_day)
    ws_, we_ = week_bounds(wk)
    wk_dates = sorted(d for d in cur_dates if ws_ <= d <= we_)
    cur_month = (last_day.year, last_day.month)
    n = len(months)
    MTD, WEEK = n, n + 1

    data = collections.defaultdict(lambda: [[0.0] * 6 for _ in range(n + 2)])
    for (nm, d), v in hist.items():
        if (d.year, d.month) in months:
            b = data[nm][months.index((d.year, d.month))]
            for i in range(6):
                b[i] += v[i]
    for (nm, d), v in cur.items():
        targets = ([MTD] if (d.year, d.month) == cur_month else []) + ([WEEK] if ws_ <= d <= we_ else [])
        for t in targets:
            for i in range(6):
                data[nm][t][i] += v[i]

    m = Masters(a.zone_master, a.item_zones, a.bom, a.categories)
    print(f"FG zones {len(m.fg_zones)} | FG BOM source {dict(collections.Counter(m.fg_source.values()))} | item zones from file {len(m.item_zones)}")
    if m.unknown_zone_names:
        print("WARNING unrecognised zone names:", sorted(m.unknown_zone_names))

    items, unzoned, nocat = [], [], 0
    for nm, (uom, typ) in meta.items():
        if typ in ("PRODUCT", "BY-PRODUCT"):
            continue
        v = data.get(nm)
        if not v or not any(abs(x) > 1e-9 for b in v for x in b):
            continue
        cat = m.category(nm, typ)
        if cat is None:
            nocat += 1
            continue
        z = m.zone(nm)
        if z is None:
            unzoned.append((None, cat, nm, uom, typ, sum(b[CONS] for b in v[:n]), sum(b[FLOOR] + b[STORE] + b[WRITEOFF] for b in v[:n])))
            continue
        items.append((z, cat, nm, uom, typ, v))

    mlabel = lambda k: dt.date(k[0], k[1], 1).strftime("%B %Y")
    periods = []
    for i, k in enumerate(months[:-1]):
        p = Period(f"m{i}", mlabel(k), i)
        p.collapsed = True
        periods.append(p)
    sum_label = f"{mlabel(months[0])} - {mlabel(months[-1])}" if n > 1 else mlabel(months[0])
    periods.append(Period("sum", sum_label))
    periods.append(Period(f"m{n - 1}", mlabel(months[-1]), n - 1))
    mtd_label = f"{dt.date(*cur_month, 1).strftime('%B %Y')} ({dt.date(*cur_month, 1).strftime('%b')} 1 to {last_day.day})"
    periods.append(Period("mtd", mtd_label, MTD))
    wk_label = f"Week {wk} ({wk_dates[0].strftime('%d-%b')} to {we_.strftime('%d-%b-%Y')})"
    periods.append(Period("wk", wk_label, WEEK))
    tail = f" {dt.date(*months[0], 1).strftime('%b %Y')} to {dt.date(*months[-1], 1).strftime('%b %Y')} + {mtd_label} & {wk_label} - Zone wise"

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    collapse = not a.expand_rows
    _, c1 = build_sheet(wb, "Wastage Summary", "Wastage Summary" + tail, "waste",
                        [i for i in items if i[1] not in BLEND_SHEETS], periods, m.zone_desc, collapse)
    counts = {"Wastage Summary": c1}
    for cat, sheet in BLEND_SHEETS.items():
        _, c = build_sheet(wb, sheet, f"{cat} Wastage Summary" + tail, "blend",
                           [i for i in items if i[1] == cat], periods, m.zone_desc, collapse)
        counts[sheet] = c
    if a.unzoned_sheet and unzoned:
        add_unzoned_sheet(wb, sorted(unzoned, key=lambda t: (t[1], t[2])))
    wb.save(a.out)

    for s, c in counts.items():
        print(f"{s}: {sum(c.values())} items  " + ", ".join(f"Zone {z}: {k}" for z, k in sorted(c.items())))
    print(f"summary {mlabel(months[0])} .. {mlabel(months[-1])} | {mtd_label} | {wk_label}")
    print(f"items without zone: {len(unzoned)}" + (f" -> {[u[2] for u in unzoned][:20]}" if unzoned else ""))
    print(f"items skipped (no category, e.g. spares/tools): {nocat}")
    print(f"saved {a.out}")


if __name__ == "__main__":
    main()
