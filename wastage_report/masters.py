"""
Master data: item -> zone and item -> category.

Zone of a raw material, in order of priority:
  1. the item-level zone master (Item Name | Zone), e.g. your filled "Items without Zone" sheet
     or the exported Zone_Master_Items file; the zone may be a number (1-9) or a zone name
     such as "Tea Bags Flavor/Herbs";
  2. otherwise via the BOM: the item -> the FGs that use it (Exploded BOM, then the BOM grid,
     then FG-name variants such as "500 GMS PKT" vs "24/500 GMS PKT") -> the FG zones from the
     FG zone master; the zone with the most FGs wins (ties -> lower zone number).
"""
import collections
import re

import openpyxl

from .utils import norm, norm_disp


# ---------------------------------------------------------------- helpers
def _find_header(ws, wanted, max_rows=8):
    """Return (row_index, {header: col}) for the first row containing all wanted headers (case-insensitive)."""
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=max_rows, values_only=True), 1):
        cells = {norm(v): j for j, v in enumerate(row) if v is not None}
        if all(any(w in k for k in cells) for w in wanted):
            return i, cells
    return None, None


def _col(cells, *names):
    for n in names:
        for k, j in cells.items():
            if k == n:
                return j
    for n in names:
        for k, j in cells.items():
            if k.startswith(n):
                return j
    return None


# ---------------------------------------------------------------- FG zone master
def load_fg_zones(path):
    """Zone master workbook: 'Active FG-12 Months' (FG name col A, Zone col E) + 'Master' (zone types)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    fg = {}
    ws = wb["Active FG-12 Months"] if "Active FG-12 Months" in wb.sheetnames else wb.worksheets[0]
    for r in ws.iter_rows(values_only=True):
        if r and r[0] and len(r) > 4 and isinstance(r[4], (int, float)):
            fg[norm(r[0])] = int(r[4])
    desc = {}
    if "Master" in wb.sheetnames:
        for r in wb["Master"].iter_rows(values_only=True):
            if r and len(r) > 4 and isinstance(r[4], (int, float)) and r[2]:
                desc[int(r[4])] = f"{norm_disp(r[2])} - {norm_disp(r[3])}"
    return fg, desc


def _zone_key(x):
    return re.sub(r"[^A-Z0-9/]", "", norm(x))


def load_item_zones(paths, zone_desc):
    """Item-level zones from any sheet that has an 'Item Name' column and a column starting with 'Zone'."""
    by_name = {_zone_key(v): k for k, v in zone_desc.items()}        # "Tea Bags - Flavor/Herbs" -> 4
    out, unknown = {}, set()
    for p in paths or []:
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        for ws in wb.worksheets:
            hr, cells = _find_header(ws, ["ITEM NAME", "ZONE"])
            if not hr:
                continue
            ci, cz = _col(cells, "ITEM NAME"), _col(cells, "ZONE")
            for r in ws.iter_rows(min_row=hr + 1, values_only=True):
                if not r or ci >= len(r) or not r[ci] or cz >= len(r) or r[cz] in (None, ""):
                    continue
                if norm(r[ci]) in ("TOTAL", "ZONE TOTAL", "GRAND TOTAL") or norm(r[cz]) == "ALL ZONES":
                    continue
                v = r[cz]
                z = int(v) if isinstance(v, (int, float)) or str(v).strip().isdigit() else by_name.get(_zone_key(v))
                if z is None:
                    m = re.match(r"ZONE\s*(\d+)", norm(v))
                    z = int(m.group(1)) if m else None
                if z:
                    out[norm(r[ci])] = z
                else:
                    unknown.add(str(v))
    return out, unknown


# ---------------------------------------------------------------- BOM
class Bom:
    def __init__(self, paths):
        self.exploded = {}                      # FG -> set(item display names); later files override
        self.item_type = {}                     # item -> ITEM TYPE (category) where the BOM has it
        self.grid_comp = collections.defaultdict(dict)
        self.grid_type = {}
        for p in paths or []:
            wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
            for ws in wb.worksheets:
                hr, cells = _find_header(ws, ["FG NAME", "ITEM NAME"], 3)
                if hr:
                    self._load_exploded(ws, hr, cells)
                    continue
                hr, cells = _find_header(ws, ["BOM HDR ITEM NAME", "ITEM_NAME", "BOM VER ITEM TYPE"], 8)
                if hr:
                    self._load_grid(ws, hr, cells)

    def _load_exploded(self, ws, hr, cells):
        cf, ci, ct = _col(cells, "FG NAME"), _col(cells, "ITEM NAME"), _col(cells, "ITEM TYPE")
        new = collections.defaultdict(set)
        for r in ws.iter_rows(min_row=hr + 1, values_only=True):
            if r and len(r) > max(cf, ci) and r[cf] and r[ci]:
                new[norm(r[cf])].add(norm_disp(r[ci]))
                if ct is not None and ct < len(r) and r[ct]:
                    self.item_type.setdefault(norm(r[ci]), norm_disp(r[ct]))
        self.exploded.update(new)

    def _load_grid(self, ws, hr, cells):
        ch, cty = _col(cells, "BOM HDR ITEM NAME"), _col(cells, "ITEM TYPE")
        cd, ci, cv = _col(cells, "IS DEFAULT"), _col(cells, "ITEM_NAME"), _col(cells, "BOM VER ITEM TYPE")
        for r in ws.iter_rows(min_row=hr + 1, values_only=True):
            if not r or len(r) <= max(ch, ci, cv) or not r[ch]:
                continue
            h = norm(r[ch])
            if cty is not None:
                self.grid_type[h] = norm(r[cty])
            if r[ci] and norm(r[ci]) != h and (cd is None or norm(r[cd]) == "Y"):
                self.grid_comp[h][norm_disp(r[ci])] = norm(r[cv])

    def grid_raw(self, h, depth=0, seen=None):
        """Raw materials of a BOM-grid header, following PRODUCT children; by-products excluded."""
        seen = seen if seen is not None else set()
        out = set()
        for it, t in self.grid_comp.get(h, {}).items():
            if t == "PRODUCT":
                if depth < 4 and norm(it) not in seen:
                    seen.add(norm(it))
                    out |= self.grid_raw(norm(it), depth + 1, seen)
            elif t != "BY-PRODUCT":
                out.add(it)
        return out


def _variant_key(x):
    x = re.sub(r"\b\d+\s*/\s*", "", norm(x))           # "24/500 GMS" -> "500 GMS"
    x = re.sub(r"[-\s]*100\s*%\s*UTZ", "", x)           # "-100% UTZ" suffix
    return re.sub(r"\s+", " ", x).strip(" -")


def bom_item_zones(bom, fg_zones):
    """Primary zone of every raw material reachable from a zoned FG; also returns per-FG source."""
    variants = collections.defaultdict(set)
    for f in bom.exploded:
        variants[_variant_key(f)].add(f)
    for f, t in bom.grid_type.items():
        if t == "PRODUCT":
            variants[_variant_key(f)].add(f)
    votes = collections.defaultdict(collections.Counter)
    source = {}
    for fg, z in fg_zones.items():
        items, src = bom.exploded.get(fg), "Exploded BOM"
        if not items:
            items, src = bom.grid_raw(fg), "BOM grid"
        if not items:
            cand = variants.get(_variant_key(fg), set()) - {fg}
            items = set().union(*[bom.exploded.get(c) or bom.grid_raw(c) for c in cand]) if cand else set()
            src = "Name variant" if items else "Not found in BOM"
        source[fg] = src
        for it in items:
            votes[norm(it)][z] += 1
    primary = {it: min(c, key=lambda z: (-c[z], z)) for it, c in votes.items() if c}
    return primary, source


# ---------------------------------------------------------------- categories
def load_categories(paths):
    """Item -> category from a previous report (row 2: 'Item Category' / 'Item Name') or a two-column list."""
    out = {}
    for p in paths or []:
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        for ws in wb.worksheets:
            hr, cells = _find_header(ws, ["ITEM NAME", "ITEM CATEGORY"])
            if not hr:
                continue
            ci, cc = _col(cells, "ITEM NAME"), _col(cells, "ITEM CATEGORY")
            for r in ws.iter_rows(min_row=hr + 1, values_only=True):
                if r and len(r) > max(ci, cc) and r[ci] and r[cc] and norm(r[ci]) not in ("TOTAL", "ZONE TOTAL", "GRAND TOTAL"):
                    out.setdefault(norm(r[ci]), norm_disp(r[cc]))
    return out


class Masters:
    def __init__(self, fg_zone_master, item_zone_files, bom_files, category_files):
        self.fg_zones, self.zone_desc = load_fg_zones(fg_zone_master)
        self.item_zones, self.unknown_zone_names = load_item_zones(item_zone_files, self.zone_desc)
        self.bom = Bom(bom_files)
        self.bom_zones, self.fg_source = bom_item_zones(self.bom, self.fg_zones)
        self.categories = load_categories(category_files)
        self.known_categories = set(self.categories.values()) | set(self.bom.item_type.values())

    def zone(self, item):
        k = norm(item)
        return self.item_zones.get(k) or self.bom_zones.get(k)

    def category(self, item, item_type=""):
        k = norm(item)
        if k in self.categories:
            return self.categories[k]
        if k in self.bom.item_type:
            return self.bom.item_type[k]
        first = k.split()[0] if k else ""
        if first in self.known_categories:
            return first
        if str(item_type).upper() == "BLEND":
            return "BLENDS"                     # uncategorised blends go to Speciality Tea
        return None
