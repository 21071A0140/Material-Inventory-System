"""
RULE BOOK CONSOLIDATOR — turn raw schedule output into structured rules.

The vision reader preserves architect-specific wording verbatim. Now we
normalize each schedule into machine-usable rules:

  HEADER SCHEDULE rows  -> {mark: {lumber_size, ply_count, material, ...}}
  SHEAR WALL SCHEDULE   -> {mark/floor: {sheathing, nail, oc_edge, oc_field, ...}}
  GENERAL NOTES         -> {stud_size_default, oc_default, sheathing_thickness, ...}

Generalization principle:
  Every architect uses the SAME LANGUAGE for wood framing — the differences
  are only in column NAMES and ARRANGEMENT. So we normalize by SEMANTICS not
  by header text. We look for "MARK" or "TAG" or "ID" -> mark. We look for
  any cell containing "x" between two numbers -> a lumber dimension. Any
  number followed by '"' or "OC" -> a spacing. This way the same code works
  whether architect A calls a column "HEADER SIZE" or architect B calls it
  "BEAM DIMENSIONS".
"""
import re
from collections import defaultdict

# ── Semantic regex over CELL CONTENT (architect-independent) ─────────────
RX_LUMBER_2X   = re.compile(r"\(?(\d)\)?\s*(\d{1,2})[xX](\d{1,2})\b")  # "(2) 2x10"
# Engineered wood like 3-1/2"x11-7/8", 5-1/4"x14", 7"x9-1/4". Allow whole numbers
# OR compound fractions (e.g. "3-1/2", "11-7/8") OR plain fractions ("7/8").
_FRAC = r'\d+(?:-\d+/\d+)?(?:/\d+)?'
RX_LUMBER_EWP  = re.compile(rf'({_FRAC})"?\s*[xX]\s*({_FRAC})"?')
RX_PLY_COUNT   = re.compile(r"\((\d)\)")
RX_MATERIAL    = re.compile(r"\b(PSL|LVL|LSL|GLULAM|SYP|SPF|DF|PT|PLY|OSB|PLYWOOD)\b", re.I)
RX_THICKNESS   = re.compile(r"(\d+/\d+)\"|^(\d+(?:\.\d+)?)\"")
RX_OC          = re.compile(r"(\d{1,2})\"\s*OC", re.I)
RX_NAIL        = re.compile(r"0\.\d{3}\"\s*x\s*\d", re.I)
RX_HEADER_MARK = re.compile(r"^H\d+-\d+[A-Z]?$")
RX_FOOTING_MK  = re.compile(r"^W?F-\d+(?:\.\d)?(?:x\d+(?:\.\d)?)?$")
RX_SHEAR_MK    = re.compile(r"^(SW-?\d+|TYPE-?\d+|[A-F])$")


def _find_column(columns, *aliases):
    """Find a column name matching any alias (case-insensitive substring)."""
    for col in columns:
        cl = col.lower()
        if any(a.lower() in cl for a in aliases):
            return col
    return None


def _parse_header_size(cell):
    """Parse '(2) 2x10' / '3-1/2"x11-7/8" PSL' / '(3) 1-3/4"x20" LVL' etc.
    Returns {lumber_size, ply_count, material} with whatever could be parsed."""
    out = {"raw": cell}
    if not cell or cell.strip() in ("-", "—"):
        return None
    s = cell.strip()

    # Detect engineered wood by material keyword first — these always use
    # compound-fraction sizes ('3-1/2"x11-7/8"'), so we must match against the
    # compound-fraction regex, not the simple 2x regex which would mis-grab.
    mat = RX_MATERIAL.search(s)
    is_ewp = bool(mat and mat.group(1).upper() in ("PSL", "LVL", "LSL", "GLULAM"))

    if is_ewp:
        m2 = RX_LUMBER_EWP.search(s)
        if m2:
            out["lumber_size"] = f'{m2.group(1)}"x{m2.group(2)}"'
            mp = RX_PLY_COUNT.search(s)
            if mp: out["ply_count"] = int(mp.group(1))
    else:
        # Try 2x style: "(2) 2x10"
        m = RX_LUMBER_2X.search(s)
        if m:
            out["ply_count"] = int(m.group(1)) if m.group(1) else 1
            out["lumber_size"] = f"{m.group(2)}x{m.group(3)}"
        else:
            m2 = RX_LUMBER_EWP.search(s)
            if m2:
                out["lumber_size"] = f'{m2.group(1)}"x{m2.group(2)}"'
                mp = RX_PLY_COUNT.search(s)
                if mp: out["ply_count"] = int(mp.group(1))

    if mat: out["material"] = mat.group(1).upper()

    return out if any(k for k in out if k != "raw") else {"raw": cell}


def normalize_header_schedule(schedule):
    """Turn a HEADER/BEAM schedule into {mark: {lumber_size, ply_count, material, raw}}."""
    cols = schedule.get("columns", [])
    mark_col   = _find_column(cols, "mark", "tag", "id", "type")
    size_col   = _find_column(cols, "size", "dimension", "section")
    flitch_col = _find_column(cols, "flitch", "plate")
    ply_col    = _find_column(cols, "ply", "plies", "number")

    out = {}
    for row in schedule.get("rows", []):
        if not row: continue
        mark = (row.get(mark_col) or "").strip() if mark_col else None
        if not mark: continue
        parsed = _parse_header_size(row.get(size_col, "")) if size_col else {}
        if not parsed: parsed = {"raw": row.get(size_col, "")}

        if ply_col and row.get(ply_col):
            try: parsed["ply_count"] = int(re.search(r"\d", row[ply_col]).group())
            except Exception: pass
        if flitch_col and row.get(flitch_col) and row[flitch_col] not in ("-", "—", None):
            parsed["flitch_plate"] = row[flitch_col]
        out[mark] = parsed
    return out


def normalize_footing_schedule(schedule):
    """Footings: {mark: {size, reinforcing_bottom, reinforcing_top, remarks, raw}}."""
    cols = schedule.get("columns", [])
    mark_col = _find_column(cols, "mark", "tag", "id")
    size_col = _find_column(cols, "size", "dimension", "width")
    bot_col  = _find_column(cols, "bottom", "bot")
    top_col  = _find_column(cols, "top")
    rem_col  = _find_column(cols, "remark", "note")

    out = {}
    for row in schedule.get("rows", []):
        mark = (row.get(mark_col) or "").strip() if mark_col else None
        if not mark: continue
        out[mark] = {
            "size": row.get(size_col),
            "reinforcing_bottom": row.get(bot_col),
            "reinforcing_top": row.get(top_col),
            "remarks": row.get(rem_col),
        }
    return out


def normalize_shear_wall_schedule(schedule):
    """Shear-wall: keyed by floor or mark depending on layout."""
    cols = schedule.get("columns", [])
    mark_col   = _find_column(cols, "mark", "type")
    floor_col  = _find_column(cols, "floor", "level", "story")
    sheath_col = _find_column(cols, "sheath", "panel")
    nail_col   = _find_column(cols, "nail", "fastener")
    edge_col   = _find_column(cols, "edge")
    field_col  = _find_column(cols, "field")
    plf_col    = _find_column(cols, "shear", "plf", "load")
    stud_col   = _find_column(cols, "end stud", "stud size", "wall end")

    out = {}
    for i, row in enumerate(schedule.get("rows", [])):
        key = (row.get(mark_col) or row.get(floor_col) or f"row_{i+1}").strip()
        entry = {
            "sheathing":  row.get(sheath_col),
            "nail":       row.get(nail_col),
            "oc_edge":    row.get(edge_col),
            "oc_field":   row.get(field_col),
            "shear_plf":  row.get(plf_col),
            "end_stud":   row.get(stud_col),
        }
        # Parse OC values if architect put '"6" OC"' style strings
        for k in ("oc_edge", "oc_field"):
            v = entry.get(k)
            if v and isinstance(v, str):
                m = RX_OC.search(v) or re.search(r'(\d+)"', v)
                if m: entry[k] = int(m.group(1))
        out[key] = entry
    return out


def normalize_window_schedule(schedule):
    """Windows: {mark: {width, height, type, qty, ...}}"""
    cols = schedule.get("columns", [])
    mark_col   = _find_column(cols, "mark", "tag", "type")
    width_col  = _find_column(cols, "width")
    height_col = _find_column(cols, "height")
    qty_col    = _find_column(cols, "qty", "quantity", "count")
    type_col   = _find_column(cols, "type", "description")

    out = {}
    for row in schedule.get("rows", []):
        mark = (row.get(mark_col) or "").strip() if mark_col else None
        if not mark: continue
        out[mark] = {
            "width":  row.get(width_col),
            "height": row.get(height_col),
            "qty":    row.get(qty_col),
            "type":   row.get(type_col),
        }
    return out


def normalize_door_schedule(schedule):
    """Doors: {mark: {width, height, type, material, qty, ...}}"""
    cols = schedule.get("columns", [])
    mark_col   = _find_column(cols, "mark", "tag", "type")
    width_col  = _find_column(cols, "width")
    height_col = _find_column(cols, "height")
    mat_col    = _find_column(cols, "material", "frame")
    type_col   = _find_column(cols, "type", "description")
    qty_col    = _find_column(cols, "qty", "quantity")

    out = {}
    for row in schedule.get("rows", []):
        mark = (row.get(mark_col) or "").strip() if mark_col else None
        if not mark: continue
        out[mark] = {
            "width":    row.get(width_col),
            "height":   row.get(height_col),
            "material": row.get(mat_col),
            "type":     row.get(type_col),
            "qty":      row.get(qty_col),
        }
    return out


# ── Schedule type classifier (architect-agnostic) ────────────────────────
def classify_schedule(schedule):
    """Decide which kind of schedule this is from its name + column names."""
    name = (schedule.get("schedule_name", "") or "").upper()
    cols_blob = " ".join((schedule.get("columns") or [])).upper()
    blob = name + " " + cols_blob
    if "HEADER" in blob or "BEAM" in blob or "LINTEL" in blob:        return "header"
    if "SHEAR" in blob:                                               return "shear_wall"
    if "FOOTING" in blob:                                             return "footing"
    if "WINDOW" in blob:                                              return "window"
    if "DOOR" in blob:                                                return "door"
    if "TRUSS" in blob:                                               return "truss"
    if "ASSEMBLY" in blob or "FASTEN" in blob or "MULTI-PLY" in blob: return "assembly_pattern"
    return "unknown"


# ── Public API ──────────────────────────────────────────────────────────
def build_rule_book(vision_results):
    """Consolidate the per-page vision outputs into one normalized rule book.

    Input:  list of {page, sheet_id, schedules: [...], notes: [...], ...}
    Output: {
              "headers":    {mark: {lumber_size, ply_count, material, ...}},
              "footings":   {mark: {size, reinforcing_bottom, ...}},
              "shear_walls":{key:  {sheathing, nail, oc_edge, oc_field, ...}},
              "windows":    {mark: {...}},  "doors": {mark: {...}},
              "notes":      [...],
              "unknown_schedules": [...] (raw, so nothing is lost),
              "_source":    [{sheet_id, page, schedule_name, type}]
            }
    """
    book = {"headers": {}, "footings": {}, "shear_walls": {},
            "windows": {}, "doors": {}, "notes": [],
            "unknown_schedules": [], "_source": []}

    for r in vision_results:
        sheet_id = r.get("sheet_id")
        page     = r.get("page")
        book["notes"].extend(r.get("notes") or [])
        for sch in r.get("schedules") or []:
            stype = classify_schedule(sch)
            book["_source"].append({
                "sheet_id": sheet_id, "page": page,
                "schedule_name": sch.get("schedule_name"),
                "type": stype, "row_count": len(sch.get("rows") or []),
            })
            if   stype == "header":     book["headers"].update(normalize_header_schedule(sch))
            elif stype == "footing":    book["footings"].update(normalize_footing_schedule(sch))
            elif stype == "shear_wall": book["shear_walls"].update(normalize_shear_wall_schedule(sch))
            elif stype == "window":     book["windows"].update(normalize_window_schedule(sch))
            elif stype == "door":       book["doors"].update(normalize_door_schedule(sch))
            else:                       book["unknown_schedules"].append(sch)
    return book


if __name__ == "__main__":
    import sys, json
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).parent))
    from extractor import scan_set
    from vision_reader import read_schedule_page_offline

    pdf = "/mnt/user-data/uploads/North_Building_Rev1_IFC_-_Combined_Set.pdf"
    scan = scan_set(pdf)
    vision_results = []
    for s in scan["sheets"]:
        if s["role"].startswith("schedule"):
            r = read_schedule_page_offline(pdf, s["page"])
            r["sheet_id"] = s["sheet_id"]
            vision_results.append(r)

    book = build_rule_book(vision_results)

    print("=" * 70)
    print("RULE BOOK (normalized from architect's schedules)")
    print("=" * 70)
    print(f"\nHEADER MARKS ({len(book['headers'])} entries):")
    for mark, info in list(book["headers"].items())[:10]:
        print(f"  {mark:8s}  {info}")
    print(f"  ... +{max(0, len(book['headers']) - 10)} more")

    print(f"\nFOOTING MARKS ({len(book['footings'])} entries):")
    for mark, info in book["footings"].items():
        print(f"  {mark:10s}  size={info['size']}  rebar_bot={info['reinforcing_bottom']}")

    print(f"\nGENERAL NOTES ({len(book['notes'])}):")
    for n in book["notes"][:5]:
        print(f"  • {n[:90]}")

    print(f"\nSOURCE TRACEABILITY:")
    for src in book["_source"]:
        print(f"  {src['sheet_id']:8s} p{src['page']}  {src['type']:18s}  {src['schedule_name']:30s}  rows={src['row_count']}")
