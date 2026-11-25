# Weekly Wastage Report

Generates the zone-wise PM / blend wastage workbook from ERP **Transaction Report** exports:

| Sheet | Items |
|---|---|
| **Wastage Summary** | PM-Specific / PM-Generic / Promo items (all categories except blends) |
| **Black Tea** | BLENDB blends, with By-product and Reclaim columns |
| **Speciality Tea** | BLENDS blends, with By-product and Reclaim columns |

Rows are grouped **Zone (1–9) → Item Category** with category *Total*, *Zone Total* (gold) and
*Grand Total* (orange) rows. The workbook opens collapsed to Zone Totals (use the outline buttons
to expand).

Column blocks, left to right:

1. each month of the summary period except the last (collapsed)
2. the summary period total (e.g. *October 2025 – October 2026*)
3. the last month of the period (e.g. *October 2026*)
4. the current month to date (e.g. *November 2026 (Nov 1 to 21)*)
5. the latest week (e.g. *Week 47 (16-Nov to 21-Nov-2026)*)

Inside the visible blocks only Consumption and the total wastage columns are shown; the detail
columns are grouped and can be expanded.

## Calculations

From each Transaction Report (Transaction types **PROCESS, WASTAGE, WRITEOFF**; Movement types
**STOCKOUT** and **STOCKIN**):

| Column | Source |
|---|---|
| Consumption | PROCESS / STOCKOUT |
| Floor Wastage | WASTAGE / STOCKOUT from any warehouse except the PM stores |
| Store Wastage | WASTAGE / STOCKOUT from WH-PM-MDKCBE or WH-PM-MDKCBE2 |
| Writeoff | WRITEOFF / STOCKOUT |
| By Prod Cons Qty (blends) | PROCESS / STOCKOUT flagged *Is By Product Consumption = Y* |
| Reclaimed Qty (blends) | PROCESS / STOCKIN on work centre *RECLAIM* |

Only the STOCKOUT leg of WASTAGE / WRITEOFF is used (the STOCKIN leg is the same quantity moving
into the wastage warehouse). STOCKIN is needed only for Reclaimed Qty.

* Total Wastage = Floor + Store + Writeoff
* Wastage before Reclaim = Floor + Store + By Prod Cons Qty
* By Prod after Reclaim = By Prod Cons Qty − Reclaimed Qty
* Wastage after Reclaim = Floor + Store + By Prod after Reclaim
* **Every %** = wastage ÷ (wastage + consumption); **0 %** when consumption ≤ 0 or wastage ≤ 0, so
  it never exceeds 100 %. Total rows use the same formula on their summed quantities (not an
  average). Values ≥ 2 % are highlighted red.

Weeks run Sunday–Saturday; Week 1 is the week of 28-Dec-2025 – 03-Jan-2026 (same numbering as the
PM report app). Product and By-Product items are excluded; items without a category (spares,
tools …) are skipped.

## Zones and categories

**Zone of an item**, in order:

1. `--item-zones` files: any sheet with an *Item Name* column and a column starting with *Zone*.
   The zone may be a number (`4`) or a zone name (`Tea Bags Flavor/Herbs`). Use the exported item
   zone master and/or a filled *Items without Zone* sheet.
2. Otherwise via the BOM: the item → the FGs that use it (Exploded BOM, then BOM grid, then FG name
   variants such as `500 GMS PKT` ↔ `24/500 GMS PKT`) → the FG's zone from `--zone-master`
   (sheet *Active FG-12 Months*, zone in column E; zone names from sheet *Master*). The zone with
   the most FGs wins; ties go to the lower zone number.

Items that still have no zone are left out; add `--unzoned-sheet` to list them in an
*Items without Zone* sheet, fill in column A and pass that file back with `--item-zones`.

**Category of an item**: `--categories` (a previous report, or an *Item Name | Item Category*
list), else the *ITEM TYPE* column of the Exploded BOM, else the first word of the item name when
it is a known category; uncategorised blends go to BLENDS.

## Usage

```bash
pip install -r requirements.txt

python -m wastage_report \
  --history  input/Transaction_Report_Oct2025-Oct2026.xlsb \
  --current  input/TR_02Nov-07Nov.xlsx input/TR_WEEK_46.xlsx input/TR_WEEK_47.xlsx \
  --zone-master input/Item_Master_Zone_wise.xlsx \
  --item-zones  input/Zone_Master_Items_9Zones.xlsx input/Items_without_Zone_filled.xlsx \
  --bom input/Exploded_BOM.xlsx input/Exploded_BOM_Supplemental.xlsx input/Additional_bom.xlsx \
  --categories  input/previous_report.xlsx \
  --out output/Wastage_Summary_Week47.xlsx
```

| Option | Meaning |
|---|---|
| `--history` | Transaction Report(s) covering the summary months |
| `--current` | this month's weekly Transaction Reports; the latest week becomes the week block, all of them the month-to-date block |
| `--summary-start`, `--summary-end` | override the summary months (`YYYY-MM`); default = the months found in `--history` |
| `--bom` | list oldest first – a later file replaces an FG's BOM from an earlier one |
| `--unzoned-sheet` | add the *Items without Zone* sheet |
| `--expand-rows` | open with item rows expanded |

**Each week:** export that week's Transaction Report (Sun–Sat, STOCKOUT + STOCKIN), add it to
`--current` and run again. When a month closes, export it into the history file (or add the
month's report to `--history`) and start a new `--current` list.

Input files are never committed (see `.gitignore`).
