"""
takeoff_parser.py — MATINV Phase 2
Reads a PlanSwift Takeoff backup ZIP (folder-per-item + Data.xml per section)
and produces a clean, wood-only measurement table.

Backup format:
  Takeoff.zip
    Takeoff/
      {LevelCode}-{ItemName}/
        Data.xml          ← item metadata (Class, Name, Wall Height, Scale Units)
        Section/Data.xml  ← first segment (DigitizerData points)
        Section1/Data.xml
        ...

Scale note: DigitizerData is in PDF coordinate space (points at 72 dpi).
The actual scale (e.g. 1/4"=1') is stored in the SwiftJob page metadata,
NOT in this backup format.  Pass scale_ft_per_pt=<value> to convert.
Common values:
  1/4"=1'  → scale = 1/216   (1 ft = 3" = 216 pt)
  3/16"=1' → scale = 1/162
  1/8"=1'  → scale = 1/108
  1"=1'    → scale = 1/72    (full-scale)

If scale_ft_per_pt is None, raw pixel totals are returned (for manual scaling).
"""

import os
import re
import math
import zipfile
import tempfile
import shutil
import html
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import Optional
from collections import defaultdict


# ─────────────────────────────────────────────────────────
# Data classes
# ─────────────────────────────────────────────────────────

@dataclass
class TakeoffItem:
    """One PlanSwift takeoff item (folder) with aggregated measurement."""
    raw_name: str           # Full name, e.g. "L2-Exterior Wall-2X6@16\" OC"
    item_class: str         # Linear / Area / Count
    level: str              # L1, L2, … Roof, Unit A1, …
    category: str           # exterior_wall / interior_wall / shear_wall / blocking / …
    stud_size: Optional[str]   # 2x4, 2x6, 2x8, 2x10, 2x12
    oc_spacing: Optional[int]  # inches: 12, 16, 24, 8
    sw_mark: Optional[str]     # SW1, SW2, … (shear-wall schedule mark)
    height_in: Optional[float] # stud height in inches (from item name suffix -50 etc.)
    is_concrete: bool          # True → skip (attaches-to-concrete or CMU substrate)
    is_wood: bool              # True → include in wood scope
    raw_total_pt: float        # sum of section lengths/areas in PDF points
    section_count: int         # number of sub-sections measured
    page_guids: list[str]      # which drawing pages contributed
    # scaled result
    total_lf: Optional[float] = None   # for Linear items in feet
    total_sf: Optional[float] = None   # for Area items in sq feet
    total_count: Optional[int] = None  # for Count items

    @property
    def is_skip(self) -> bool:
        """True if this item should be excluded from wood scope."""
        return self.is_concrete or not self.is_wood


# ─────────────────────────────────────────────────────────
# Name parser
# ─────────────────────────────────────────────────────────

# CMU/Concrete indicators in the item name
_RE_CONCRETE = re.compile(r'[-_\s]Conc(rete)?$|[-_]Con$|-CMU|-Masonry', re.I)

# Lumber dimensions: 2x4, 2x6, 2x8, 2x10, 2x12, 4x4, 4x6, 6x6 etc.
_RE_STUD = re.compile(r'\b([2-9])[Xx]([4-9]|1[0-2])\b')

# EWP sizes: PSL, LVL, LSL, Glulam column items like "4X6 PSL"
_RE_EWP = re.compile(r'\b([3-9]|1\d)\s*[Xx]\s*([4-9]|1[0-9])\s*(PSL|LVL|LSL|Glulam|GLB)', re.I)

# OC spacing
_RE_OC = re.compile(r'@\s*(\d+)\s*"\s*OC', re.I)

# Shear-wall mark
_RE_SW = re.compile(r'\bSW\s*(\d+[A-Za-z]?)\b', re.I)

# Height suffix: trailing -50 -51 -53 -104 etc. (in inches)
_RE_HEIGHT_SUFFIX = re.compile(r'-(\d{2,3})$')

# Level prefix patterns
_RE_LEVEL = re.compile(
    r'^(L\d+|Roof(?:\s+Tower)?|Low\s+Roof|Unit\s+[A-Z0-9]+(?:\s+Alt(?:\s+\d)?)?)',
    re.I
)

# Category keywords (order matters — more specific first)
_CATEGORY_MAP = [
    ('header',        re.compile(r'\bHeader\b', re.I)),
    ('blocking',      re.compile(r'\bBlocking\b|\bBlo\b', re.I)),
    ('sheathing',     re.compile(r'\bSheathing\b', re.I)),
    ('shear_wall',    re.compile(r'\bShear\s+Wall\b', re.I)),
    ('beam',          re.compile(r'\bBeam\b|\bPSL\b|\bLVL\b|\bLSL\b|\bGlulam\b', re.I)),
    ('balcony',       re.compile(r'\bBalcony\b', re.I)),
    ('floor',         re.compile(r'\bFloor\b', re.I)),
    ('corner_stud',   re.compile(r'\bCorner\s+Stud', re.I)),
    ('end_stud',      re.compile(r'\bEnd\s+Stud', re.I)),
    ('stair_wall',    re.compile(r'\bStair\b', re.I)),
    ('elevator_wall', re.compile(r'\bElevator\b', re.I)),
    ('corridor_wall', re.compile(r'\bCorridor\b', re.I)),
    ('demising_wall', re.compile(r'\bDemising\b', re.I)),
    ('exterior_wall', re.compile(r'\bExterior\b', re.I)),
    ('interior_wall', re.compile(r'\bInterior\b', re.I)),
]

def _parse_item_name(name: str) -> dict:
    """Parse a PlanSwift item name into structured fields."""
    out = dict(
        level=None, category='other', stud_size=None,
        oc_spacing=None, sw_mark=None, height_in=None,
        is_concrete=False, is_wood=False,
    )

    # --- CMU / Concrete filter ---
    out['is_concrete'] = bool(_RE_CONCRETE.search(name))

    # --- Level ---
    m = _RE_LEVEL.match(name)
    if m:
        out['level'] = m.group(1).strip()

    rest = name

    # --- SW mark ---
    m = _RE_SW.search(rest)
    if m:
        out['sw_mark'] = f"SW{m.group(1)}"

    # --- OC spacing ---
    m = _RE_OC.search(rest)
    if m:
        out['oc_spacing'] = int(m.group(1))

    # --- EWP first (so 3-1/2"x11-7/8" PSL doesn't get grabbed by 2x regex) ---
    m = _RE_EWP.search(rest)
    if m:
        out['stud_size'] = f"{m.group(1)}x{m.group(2)} {m.group(3).upper()}"
        out['is_wood'] = True
    else:
        # --- Regular lumber ---
        m = _RE_STUD.search(rest)
        if m:
            out['stud_size'] = f"{m.group(1)}x{m.group(2)}"
            out['is_wood'] = True

    # --- Height suffix ---
    m = _RE_HEIGHT_SUFFIX.search(rest)
    if m:
        try:
            out['height_in'] = float(m.group(1))
        except ValueError:
            pass

    # --- Category ---
    for cat, pat in _CATEGORY_MAP:
        if pat.search(rest):
            out['category'] = cat
            break

    return out


# ─────────────────────────────────────────────────────────
# DigitizerData geometry
# ─────────────────────────────────────────────────────────

def _polyline_length(digi_xml: str) -> float:
    """Sum of Euclidean segment lengths for a polyline in DigitizerData XML."""
    try:
        root = ET.fromstring(digi_xml)
        pts = [(float(p.get('X', 0)), float(p.get('Y', 0)))
               for p in root.iter('Point')]
        if len(pts) < 2:
            return 0.0
        return sum(math.hypot(pts[i+1][0] - pts[i][0],
                               pts[i+1][1] - pts[i][1])
                   for i in range(len(pts) - 1))
    except Exception:
        return 0.0


def _polygon_area(digi_xml: str) -> float:
    """Shoelace area for a closed polygon in DigitizerData XML."""
    try:
        root = ET.fromstring(digi_xml)
        pts = [(float(p.get('X', 0)), float(p.get('Y', 0)))
               for p in root.iter('Point')]
        n = len(pts)
        if n < 3:
            return 0.0
        return abs(sum(pts[i][0] * (pts[(i+1) % n][1] - pts[(i-1) % n][1])
                       for i in range(n))) / 2.0
    except Exception:
        return 0.0


# ─────────────────────────────────────────────────────────
# Section reader
# ─────────────────────────────────────────────────────────

def _read_section(data_xml_path: str, item_class: str) -> tuple[float, str]:
    """
    Read one Section/Data.xml.  Returns (measurement_in_pt, page_guid).
    measurement is polyline length for Linear, polygon area for Area,
    0 for Count (counts are in points metadata, not geometry).
    """
    try:
        root = ET.parse(data_xml_path).getroot()
        page_guid = ''
        digi_text = None

        for p in root.iter('Property'):
            pname = p.get('Name', '')
            if pname == 'PageGUID' and p.text:
                page_guid = p.text.strip()
            elif pname == 'DigitizerData' and p.text:
                digi_text = html.unescape(p.text)

        if not digi_text:
            return 0.0, page_guid

        if item_class == 'Linear':
            meas = _polyline_length(digi_text)
        elif item_class == 'Area':
            meas = _polygon_area(digi_text)
        else:  # Count
            meas = 0.0  # counts are not geometric sums

        return meas, page_guid
    except Exception:
        return 0.0, ''


# ─────────────────────────────────────────────────────────
# Item reader
# ─────────────────────────────────────────────────────────

def _read_item(item_dir: str) -> Optional[TakeoffItem]:
    """Read one takeoff item folder. Returns None if not a framing item."""
    data_xml = os.path.join(item_dir, 'Data.xml')
    if not os.path.exists(data_xml):
        return None

    try:
        root = ET.parse(data_xml).getroot()
    except Exception:
        return None

    item_class = root.get('Class', '')
    if item_class not in ('Linear', 'Area', 'Count'):
        return None

    full_name = root.get('Name', os.path.basename(item_dir))

    # Parse name
    parsed = _parse_item_name(full_name)

    # Read all sections
    total_pt = 0.0
    section_count = 0
    page_guids: list[str] = []

    for entry in sorted(os.listdir(item_dir)):
        entry_path = os.path.join(item_dir, entry)
        if not os.path.isdir(entry_path):
            continue
        sec_xml = os.path.join(entry_path, 'Data.xml')
        if not os.path.exists(sec_xml):
            continue

        # Recurse for nested "Subtract Section" etc.
        meas, pguid = _read_section(sec_xml, item_class)
        if meas > 0 or pguid:
            total_pt += meas
            section_count += 1
            if pguid and pguid not in page_guids:
                page_guids.append(pguid)

    return TakeoffItem(
        raw_name=full_name,
        item_class=item_class,
        level=parsed['level'] or 'Unknown',
        category=parsed['category'],
        stud_size=parsed['stud_size'],
        oc_spacing=parsed['oc_spacing'],
        sw_mark=parsed['sw_mark'],
        height_in=parsed['height_in'],
        is_concrete=parsed['is_concrete'],
        is_wood=parsed['is_wood'],
        raw_total_pt=total_pt,
        section_count=section_count,
        page_guids=page_guids,
    )


# ─────────────────────────────────────────────────────────
# Scale application
# ─────────────────────────────────────────────────────────

# Standard PlanSwift scale factors (ft per PDF point)
# PlanSwift coordinate units are PDF points (1 pt = 1/72 inch)
SCALE_PRESETS = {
    '1/4"=1\'-0"':  1 / 216,   # 1 ft = 3" on paper = 3*72 pt
    '3/16"=1\'-0"': 1 / 162,
    '1/8"=1\'-0"':  1 / 108,
    '1"=1\'-0"':    1 / 72,
    '1/2"=1\'-0"':  1 / 108,   # same as 1/8 numerically? no: 1/2"=1' → 1ft=2" paper=144pt
}

# Correct values:
SCALE_PRESETS = {
    '1/4"=1\'-0"':  1 / 216,   # 1 ft on paper = 3 inches = 216 pt
    '3/16"=1\'-0"': 1 / 162,   # 1 ft = 2.25" = 162 pt
    '1/8"=1\'-0"':  1 / 108,   # 1 ft = 1.5" = 108 pt
    '1/2"=1\'-0"':  1 / 432,   # 1 ft = 6" = 432 pt — wait, larger paper
    '1"=1\'-0"':    1 / 864,   # 1 ft = 12" = 864 pt
}
# Actually: scale S means S inches on paper = 1 foot real
# 1/4"=1' means 0.25" on paper = 1 ft real → 1 ft real = 0.25 * 72 = 18 pt
# Let me recalculate:
SCALE_PRESETS = {
    '1/4"=1\'-0"':  1 / 18,    # 0.25" * 72 pt/in = 18 pt per ft
    '3/16"=1\'-0"': 1 / 13.5,  # 0.1875 * 72
    '1/8"=1\'-0"':  1 / 9,     # 0.125 * 72
    '1/2"=1\'-0"':  1 / 36,    # 0.5 * 72
    '1"=1\'-0"':    1 / 72,    # 1 * 72
    '3/32"=1\'-0"': 1 / 6.75,
}


def apply_scale(items: list[TakeoffItem], scale_ft_per_pt: float) -> None:
    """Apply scale factor to all items (mutates in place)."""
    for item in items:
        if item.item_class == 'Linear':
            item.total_lf = item.raw_total_pt * scale_ft_per_pt
        elif item.item_class == 'Area':
            item.total_sf = item.raw_total_pt * (scale_ft_per_pt ** 2)
        # Count items stay as None (they need different treatment)


# ─────────────────────────────────────────────────────────
# Main parser
# ─────────────────────────────────────────────────────────

def parse_takeoff_zip(
    zip_path: str,
    scale_ft_per_pt: Optional[float] = None,
    wood_only: bool = True,
) -> list[TakeoffItem]:
    """
    Parse a PlanSwift Takeoff backup ZIP.

    Args:
        zip_path:        Path to Takeoff.zip
        scale_ft_per_pt: Conversion factor (feet per PDF point). 
                         Use SCALE_PRESETS['1/4"=1\'-0"'] for quarter-inch scale.
                         If None, raw_total_pt values are returned unscaled.
        wood_only:       If True (default), exclude items with is_concrete=True
                         or is_wood=False.

    Returns:
        List of TakeoffItem objects sorted by level + raw_name.
    """
    tmp_dir = tempfile.mkdtemp(prefix='takeoff_')
    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            zf.extractall(tmp_dir)

        # Find the Takeoff folder
        takeoff_root = None
        for root, dirs, files in os.walk(tmp_dir):
            if os.path.basename(root) == 'Takeoff':
                takeoff_root = root
                break

        if not takeoff_root:
            raise FileNotFoundError("Could not find 'Takeoff' folder inside ZIP")

        items: list[TakeoffItem] = []
        for entry in os.listdir(takeoff_root):
            entry_path = os.path.join(takeoff_root, entry)
            if not os.path.isdir(entry_path):
                continue
            item = _read_item(entry_path)
            if item is None:
                continue
            if wood_only and item.is_skip:
                continue
            items.append(item)

        # Sort by level numeric, then name
        def _sort_key(it):
            m = re.match(r'L(\d+)', it.level)
            lvl_n = int(m.group(1)) if m else 99
            return (lvl_n, it.level, it.raw_name)

        items.sort(key=_sort_key)

        # Apply scale if provided
        if scale_ft_per_pt is not None:
            apply_scale(items, scale_ft_per_pt)

        return items

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ─────────────────────────────────────────────────────────
# Summary helpers
# ─────────────────────────────────────────────────────────

def summarize_by_level(items: list[TakeoffItem]) -> dict:
    """
    Aggregate total LF per level per stud size.
    Returns: { level: { stud_size: total_lf } }
    """
    result = defaultdict(lambda: defaultdict(float))
    for item in items:
        if item.item_class == 'Linear' and item.stud_size and item.total_lf is not None:
            result[item.level][item.stud_size] += item.total_lf
    return {k: dict(v) for k, v in result.items()}


def to_dataframe(items: list[TakeoffItem]):
    """Convert item list to a pandas DataFrame (requires pandas)."""
    import pandas as pd
    rows = []
    for it in items:
        rows.append({
            'level': it.level,
            'raw_name': it.raw_name,
            'class': it.item_class,
            'category': it.category,
            'stud_size': it.stud_size,
            'oc_spacing': it.oc_spacing,
            'sw_mark': it.sw_mark,
            'height_in': it.height_in,
            'is_concrete': it.is_concrete,
            'raw_total_pt': it.raw_total_pt,
            'total_lf': it.total_lf,
            'total_sf': it.total_sf,
            'section_count': it.section_count,
        })
    return pd.DataFrame(rows)


# ─────────────────────────────────────────────────────────
# CLI / test
# ─────────────────────────────────────────────────────────

if __name__ == '__main__':
    import sys
    import json

    zip_path = sys.argv[1] if len(sys.argv) > 1 else '/mnt/user-data/uploads/Takeoff.zip'
    
    # Try 1/4"=1' scale as default
    scale = SCALE_PRESETS['1/4"=1\'-0"']
    print(f"Scale: 1/4\"=1'-0\"  ({scale:.6f} ft/pt)\n")

    items = parse_takeoff_zip(zip_path, scale_ft_per_pt=scale, wood_only=True)

    print(f"Loaded {len(items)} wood items\n")

    # Show by level + category
    by_level = defaultdict(lambda: defaultdict(list))
    for it in items:
        by_level[it.level][it.category].append(it)

    for level in sorted(by_level, key=lambda l: (int(re.search(r'\d+', l).group()) if re.search(r'\d+', l) else 99, l)):
        cats = by_level[level]
        total_wall_lf = sum(
            it.total_lf or 0
            for cat, cat_items in cats.items()
            for it in cat_items
            if it.item_class == 'Linear' and 'wall' in cat
        )
        print(f"\n{'='*60}")
        print(f"  {level}  —  {sum(len(v) for v in cats.values())} items  |  ~{total_wall_lf:.0f} LF walls")
        print(f"{'='*60}")
        for cat in sorted(cats):
            cat_items = cats[cat]
            lf_total = sum(it.total_lf or 0 for it in cat_items if it.item_class == 'Linear')
            sf_total = sum(it.total_sf or 0 for it in cat_items if it.item_class == 'Area')
            print(f"  {cat:20s}  {len(cat_items):3d} items  ", end='')
            if lf_total:
                print(f"  {lf_total:8.1f} LF", end='')
            if sf_total:
                print(f"  {sf_total:8.1f} SF", end='')
            print()
            for it in cat_items[:3]:
                val = it.total_lf or it.total_sf or it.raw_total_pt
                unit = 'LF' if it.total_lf else ('SF' if it.total_sf else 'pt')
                print(f"    {it.raw_name[:60]:60s}  {val:8.1f} {unit}")
            if len(cat_items) > 3:
                print(f"    ... and {len(cat_items)-3} more")

    # Summary table
    print("\n\n" + "="*60)
    print("SUMMARY: LF by level + stud size")
    print("="*60)
    summary = summarize_by_level(items)
    all_sizes = sorted(set(s for d in summary.values() for s in d))
    header = f"{'Level':10s}" + "".join(f"  {s:6s}" for s in all_sizes)
    print(header)
    print("-" * len(header))
    for lvl in sorted(summary, key=lambda l: (int(re.search(r'\d+', l).group()) if re.search(r'\d+', l) else 99, l)):
        row = f"{lvl:10s}"
        for s in all_sizes:
            v = summary[lvl].get(s, 0)
            row += f"  {v:6.0f}" if v else f"  {'':6s}"
        print(row)
