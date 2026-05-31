"""
drawing_reader.py — MATINV Phase 2: Drawing Intelligence Engine
================================================================
Reads uploaded architectural/structural PDF drawings and extracts:
  1. Sheet index (sheet ID, title, discipline, role)
  2. Stud specifications from structural general notes
  3. Header/Beam schedule
  4. Shear wall schedule
  5. Floor construction notes (truss depth)
  6. Wall type classifications from plan sheets
  7. Scale per page

Returns a structured session dict that:
  (a) Contains everything needed to run the recipe engine without Q&A
  (b) Lists any questions that need user input when data is missing/ambiguous

This module learns from multiple reference projects (Willow Way, Silver City,
future uploads) to get better over time — the rules are architect-agnostic.
"""

import re
import json
import math
import tempfile
import shutil
from pathlib import Path
from collections import defaultdict
from typing import Optional

try:
    import fitz          # PyMuPDF
    FITZ_OK = True
except ImportError:
    FITZ_OK = False

try:
    import anthropic
    CLAUDE_OK = True
except ImportError:
    CLAUDE_OK = False


# ─────────────────────────────────────────────────────────────────────────────
# Scale detection
# ─────────────────────────────────────────────────────────────────────────────

# Matches "SCALE: 1/4" = 1'-0"" or "SCALE: 1/4"=1'" etc.
_SCALE_RX = re.compile(
    r'SCALE[:\s]+(\d+(?:/\d+)?)"?\s*=\s*1\'[-–]?0?"?',
    re.I
)
# Maps "1/4" → ft_per_pt conversion (PDF pts at 72dpi)
_SCALE_TABLE = {
    "1/4":  1/18.0,   # 0.25" * 72pt/in = 18pt per ft
    "3/16": 1/13.5,
    "1/8":  1/9.0,
    "1":    1/72.0,
    "3/4":  1/54.0,
    "1/2":  1/36.0,
    "3/32": 1/6.75,
    "1/16": 1/4.5,
}

def detect_scale_on_page(page) -> Optional[float]:
    """Return scale in ft/pt for a PDF page, or None if not found."""
    text = page.get_text("text")
    m = _SCALE_RX.search(text)
    if m:
        frac = m.group(1).strip()
        return _SCALE_TABLE.get(frac)
    # Also scan word-by-word for split "1/4" + "=" + "1'"
    words = [w[4] for w in page.get_text("words")]
    blob = " ".join(words)
    m2 = _SCALE_RX.search(blob)
    if m2:
        return _SCALE_TABLE.get(m2.group(1).strip())
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Stud specification extractor (from structural general notes)
# ─────────────────────────────────────────────────────────────────────────────

# These patterns cover: "2x6 AT 12" OC", "2X6@12"OC", "(2) 2x4 AT 8" OC"
_STUD_SPEC_RX = re.compile(
    r'(?:\((\d)\)\s*)?(\d)[xX](\d{1,2})\s*(?:AT|@)\s*(\d{1,2})"\s*O\.?C',
    re.I
)
# Wall type keywords
_WALL_TYPE_RX = {
    "exterior": re.compile(r'\bEXTERIOR\b', re.I),
    "interior": re.compile(r'\bINTERIOR\b', re.I),
    "corridor": re.compile(r'\bCORRIDOR\b', re.I),
    "demising": re.compile(r'\bDEMISING\b|\bTENANT\b', re.I),
    "shear":    re.compile(r'\bSHEAR\b', re.I),
}
# Load keywords: "SUPPORTING (4) FLOORS", "SUPPORTING ROOF ONLY"
_LOAD_RX = re.compile(r'SUPPORTING\s+(?:\((\d)\)\s+FLOORS?\s+AND\s+ROOF|ROOF\s+ONLY)', re.I)

def extract_stud_specs(text: str) -> dict:
    """
    Parse the structural general notes text and return a stud-spec lookup:
    {
      "exterior": {4: ("2x6", 12), 3: ("2x6", 12), 2: ("2x6", 16), 1: ("2x6", 16), 0: ("2x6", 16)},
      "interior": {4: ("2x4", 8),  ...},
      ...
    }
    floors_supported = number of floors above (4 = bottom floor of 5-story, 0 = roof only)
    """
    specs = {}
    # Split into sections by wall type heading
    # Find all stud specs with context
    lines = text.split('\n')
    current_wall_type = None
    current_load = None

    for line in lines:
        line = line.strip()
        # Detect wall type context
        for wtype, rx in _WALL_TYPE_RX.items():
            if rx.search(line):
                current_wall_type = wtype
                break
        # Detect load context
        lm = _LOAD_RX.search(line)
        if lm:
            current_load = int(lm.group(1)) if lm.group(1) else 0
        # Detect stud spec
        for m in _STUD_SPEC_RX.finditer(line):
            ply  = int(m.group(1)) if m.group(1) else 1
            d1   = m.group(2)
            d2   = m.group(3)
            oc   = int(m.group(4))
            size = f"{d1}x{d2}"
            if current_wall_type and current_load is not None:
                wt = current_wall_type
                if wt not in specs:
                    specs[wt] = {}
                specs[wt][current_load] = (size, oc, ply)

    return specs


# ─────────────────────────────────────────────────────────────────────────────
# Truss/floor depth extractor
# ─────────────────────────────────────────────────────────────────────────────

_TRUSS_DEPTH_RX = re.compile(r'(\d{1,2})"\s*DEEP\s*(?:PRE-ENGINEERED\s+)?WOOD\s+TRUSSES?', re.I)
_TRUSS_OC_RX    = re.compile(r'TRUSSES?\s+(?:AT|@)\s*(\d{1,2})"\s*O\.?C', re.I)

def extract_truss_info(text: str) -> dict:
    """Return {depth_in, oc_in} from floor construction notes."""
    result = {}
    m = _TRUSS_DEPTH_RX.search(text)
    if m:
        result['depth_in'] = int(m.group(1))
    m2 = _TRUSS_OC_RX.search(text)
    if m2:
        result['oc_in'] = int(m2.group(1))
    return result


# ─────────────────────────────────────────────────────────────────────────────
# Floor-to-floor height extractor (from wall sections / elevations)
# ─────────────────────────────────────────────────────────────────────────────

# Matches "12'-0"", "9'-1 1/8"", "10' - 0"" etc.
_FTF_RX = re.compile(r"(\d{1,2})'\s*[-–]?\s*(\d{1,2})\s*(?:(\d+)/(\d+))?\s*\"")

def parse_ft_in(text: str) -> Optional[float]:
    """Parse a dimension like "12'-0"" or "9'-1 1/8"" into decimal inches."""
    m = _FTF_RX.search(text)
    if not m:
        return None
    ft   = int(m.group(1))
    inch = int(m.group(2))
    frac = (int(m.group(3)) / int(m.group(4))) if m.group(3) else 0.0
    return ft * 12 + inch + frac


# ─────────────────────────────────────────────────────────────────────────────
# Page classifier (reuses extractor.py logic but standalone)
# ─────────────────────────────────────────────────────────────────────────────

_SHEET_ID_RX = re.compile(r"\b([A-Z]{1,2})(\d{1,2})\.(\d{1,2}[A-Za-z]?)\b")

_ROLE_RULES = [
    ("cover",            [r"cover sheet", r"sheet index", r"index of drawings"]),
    ("general_notes",    [r"general notes", r"symbol legend", r"abbreviations"]),
    ("schedule_header",  [r"header.{0,10}schedule", r"beam schedule", r"lintel schedule"]),
    ("schedule_shear",   [r"shear wall schedule", r"shear schedule"]),
    ("schedule_footing", [r"footing schedule", r"foundation schedule"]),
    ("schedule_window",  [r"window schedule"]),
    ("schedule_door",    [r"door schedule"]),
    ("floor_plan",       [r"floor plan", r"framing plan", r"slab plan", r"foundation plan"]),
    ("roof_plan",        [r"roof plan", r"roof framing"]),
    ("wall_section",     [r"wall section", r"building section", r"typ.*section"]),
    ("elevation",        [r"\belevation\b"]),
    ("detail",           [r"\bdetail\b"]),
]

def classify_page(sheet_id: str, title: str) -> str:
    title_l = (title or "").lower()
    for role, pats in _ROLE_RULES:
        if any(re.search(p, title_l) for p in pats):
            return role
    # Fallback: use sheet ID prefix
    m = re.match(r"^([A-Z]{1,2})(\d)", sheet_id or "")
    if m:
        pfx = m.group(1)
        num = m.group(2)
        if pfx == "S" and num == "0": return "general_notes"
        if pfx in ("A", "AN") and num in ("2","3"): return "floor_plan"
        if pfx == "S" and num in ("2","3"): return "floor_plan"  # structural framing plans
    return "unknown"


def detect_sheet_id_and_title(page) -> tuple[str, str]:
    """Extract sheet ID and title from title block."""
    W, H = page.rect.width, page.rect.height
    spans = []
    for blk in page.get_text("dict")["blocks"]:
        if "lines" not in blk: continue
        for line in blk["lines"]:
            for span in line["spans"]:
                t = span["text"].strip()
                if not t: continue
                x, y = span["bbox"][0], span["bbox"][1]
                if x > W * 0.55 or y > H * 0.80 or (x > W * 0.50 and y < H * 0.12):
                    spans.append({"size": span["size"], "text": t, "x": x, "y": y})

    # Sheet ID = largest font matching pattern
    sheet_id = ""
    for s in sorted(spans, key=lambda s: -s["size"]):
        m = _SHEET_ID_RX.search(s["text"])
        if m:
            sheet_id = m.group(0)
            break

    # Title = second largest clean text
    title = ""
    boiler = re.compile(r"SCALE|LICENSE|www\.|@|\.com|PERMIT|©|\d{2,4}[-/]\d", re.I)
    for s in sorted(spans, key=lambda s: -s["size"]):
        t = s["text"]
        if len(t) < 5 or boiler.search(t): continue
        if _SHEET_ID_RX.fullmatch(t): continue
        if not any(c.isalpha() for c in t): continue
        title = t
        break

    return sheet_id, title


# ─────────────────────────────────────────────────────────────────────────────
# Vision schedule reader (uses Claude API)
# ─────────────────────────────────────────────────────────────────────────────

_SCHEDULE_PROMPT = """You are reading a construction drawing. Extract ALL schedule tables visible on this page.

For each schedule return:
{
  "schedule_name": "exact title",
  "columns": ["COL1", "COL2", ...],
  "rows": [{"COL1": "val", "COL2": "val"}, ...]
}

Also extract any GENERAL NOTES that specify stud sizes / OC spacing / lumber grade.

Return ONLY JSON:
{
  "schedules": [...],
  "stud_notes": ["text of any stud/lumber spec notes"],
  "truss_notes": ["text of any floor truss / floor construction notes"],
  "floor_heights": ["any floor-to-floor height dimensions mentioned"]
}"""

def read_page_with_vision(pdf_path: str, page_number: int, client=None) -> dict:
    """Render page and send to Claude vision. Returns structured data."""
    if not FITZ_OK:
        return {"schedules": [], "stud_notes": [], "truss_notes": [], "floor_heights": []}

    try:
        import base64
        doc  = fitz.open(pdf_path)
        page = doc[page_number - 1]
        mat  = fitz.Matrix(2.5, 2.5)
        pix  = page.get_pixmap(matrix=mat)
        img_bytes = pix.tobytes("jpeg")
        img_b64   = base64.standard_b64encode(img_bytes).decode()

        if client is None and CLAUDE_OK:
            import os
            client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

        resp = client.messages.create(
            model="claude-opus-4-5",
            max_tokens=4000,
            messages=[{"role": "user", "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": img_b64}},
                {"type": "text",  "text": _SCHEDULE_PROMPT},
            ]}]
        )
        text  = "".join(b.text for b in resp.content if b.type == "text")
        clean = re.sub(r"^```(?:json)?\s*|\s*```\s*$", "", text.strip(), flags=re.M)
        return json.loads(clean)
    except Exception as e:
        return {"schedules": [], "stud_notes": [], "truss_notes": [], "floor_heights": [], "_error": str(e)}


# ─────────────────────────────────────────────────────────────────────────────
# Main scan session
# ─────────────────────────────────────────────────────────────────────────────

def scan_drawings(pdf_paths: list[str], client=None) -> dict:
    """
    Scan a set of PDF drawings and return a session dict:
    {
      "sheets":         [{page, sheet_id, title, role, scale, pdf}],
      "stud_specs":     {wall_type: {floors_above: (size, oc, ply)}},
      "truss_info":     {depth_in, oc_in},
      "header_schedule":{mark: {lumber_size, material, ply_count, stock_ft}},
      "shear_schedule": {mark: {...}},
      "floor_heights":  {level: inches},
      "scales":         {sheet_id: ft_per_pt},
      "questions":      [{id, question, type, choices, hint}],
      "log_lines":      [{type, icon, msg}],
    }
    """
    session = {
        "sheets": [], "stud_specs": {}, "truss_info": {},
        "header_schedule": {}, "shear_schedule": {},
        "floor_heights": {}, "scales": {},
        "questions": [], "log_lines": [],
        "pages_scanned": 0, "schedules_found": 0,
        "stud_notes_found": False, "scales_detected": [],
    }
    logs = session["log_lines"]

    def log(t, icon, msg):
        logs.append({"type": t, "icon": icon, "msg": msg})

    if not FITZ_OK:
        log("err", "✗", "PyMuPDF not installed — pip install pymupdf")
        return session

    # ── Phase 1: Page scan (text extraction, no vision) ──────────────────────
    all_text_by_role: dict[str, str] = defaultdict(str)
    schedule_pages = []

    for pdf_path in pdf_paths:
        try:
            doc = fitz.open(pdf_path)
            log("info", "📄", f"{Path(pdf_path).name} — {doc.page_count} pages")

            for i in range(doc.page_count):
                page = doc[i]
                sid, title = detect_sheet_id_and_title(page)
                role  = classify_page(sid, title)
                scale = detect_scale_on_page(page)
                text  = page.get_text("text")

                sheet_rec = {
                    "page": i + 1, "sheet_id": sid, "title": title,
                    "role": role, "scale": scale, "pdf": pdf_path
                }
                session["sheets"].append(sheet_rec)
                session["pages_scanned"] += 1

                if scale:
                    session["scales"][sid or f"p{i+1}"] = scale
                    if f"{1/scale:.1f}" not in session["scales_detected"]:
                        session["scales_detected"].append(_scale_label(scale))

                all_text_by_role[role] += "\n" + text

                # Mark schedule pages for vision pass
                if role.startswith("schedule"):
                    schedule_pages.append(sheet_rec)
                    session["schedules_found"] += 1
                    log("ok", "📋", f"Schedule page: {sid or f'p{i+1}'} — {title or role}")

                # Extract notes directly from text (faster than vision)
                if role in ("general_notes",) or (sid and sid.startswith("S0")):
                    truss = extract_truss_info(text)
                    if truss.get("depth_in"):
                        session["truss_info"].update(truss)
                        log("ok", "🏗", f"Truss depth: {truss['depth_in']}\" @ {truss.get('oc_in',24)}\" OC")

                    specs = extract_stud_specs(text)
                    if specs:
                        session["stud_specs"].update(specs)
                        session["stud_notes_found"] = True
                        log("ok", "📝", f"Stud specs found on {sid or f'p{i+1}'}: {list(specs.keys())}")

        except Exception as e:
            log("err", "✗", f"Error reading {Path(pdf_path).name}: {e}")

    # ── Phase 2: Vision pass on schedule pages ────────────────────────────────
    if client and schedule_pages:
        log("info", "👁", f"Sending {len(schedule_pages)} schedule page(s) to Claude vision…")
        for sp in schedule_pages:
            try:
                vdata = read_page_with_vision(sp["pdf"], sp["page"], client=client)
                # Process header schedule
                for sch in vdata.get("schedules", []):
                    sname = (sch.get("schedule_name") or "").upper()
                    if any(k in sname for k in ("HEADER", "BEAM", "LINTEL")):
                        _parse_header_schedule(sch, session["header_schedule"])
                        log("ok", "📊", f"Header schedule: {len(session['header_schedule'])} marks from {sp['sheet_id']}")
                    elif "SHEAR" in sname:
                        _parse_shear_schedule(sch, session["shear_schedule"])
                        log("ok", "📊", f"Shear schedule: {len(session['shear_schedule'])} entries")
                # Stud notes from vision if not already found
                for note in vdata.get("stud_notes", []):
                    if not session["stud_notes_found"]:
                        specs = extract_stud_specs(note)
                        if specs:
                            session["stud_specs"].update(specs)
                            session["stud_notes_found"] = True
                # Truss notes
                for note in vdata.get("truss_notes", []):
                    t = extract_truss_info(note)
                    if t.get("depth_in"):
                        session["truss_info"].update(t)
            except Exception as e:
                log("warn", "⚠", f"Vision error on page {sp['page']}: {e}")
    elif not client:
        log("info", "ℹ", "No API key — skipping vision pass (schedules will use text extraction)")

    # ── Phase 3: Also scan general notes pages via vision if stud specs missing ─
    if client and not session["stud_notes_found"]:
        notes_pages = [s for s in session["sheets"] if s["role"] == "general_notes"][:2]
        for sp in notes_pages:
            try:
                vdata = read_page_with_vision(sp["pdf"], sp["page"], client=client)
                for note in vdata.get("stud_notes", []):
                    specs = extract_stud_specs(note)
                    if specs:
                        session["stud_specs"].update(specs)
                        session["stud_notes_found"] = True
                        log("ok", "📝", f"Stud specs (vision) from {sp['sheet_id']}")
                for note in vdata.get("truss_notes", []):
                    t = extract_truss_info(note)
                    if t.get("depth_in"):
                        session["truss_info"].update(t)
            except Exception:
                pass

    # ── Phase 4: Build questions for missing data ─────────────────────────────
    questions = []

    if not session["truss_info"].get("depth_in"):
        questions.append({
            "id": "truss_depth",
            "question": "What is the floor truss depth?",
            "hint": "Check the structural floor framing notes (e.g. S2.01). Typically 20\" or 24\".",
            "type": "choice",
            "choices": [
                {"value": "16", "label": "16\" deep"},
                {"value": "18", "label": "18\" deep"},
                {"value": "20", "label": "20\" deep"},
                {"value": "24", "label": "24\" deep"},
            ]
        })
        log("warn", "❓", "Truss depth not found — will ask user")

    if not session["stud_notes_found"]:
        questions.append({
            "id": "stud_spec_exterior",
            "question": "What stud size/spacing for exterior load-bearing walls?",
            "hint": "Check structural general notes. Usually 2x6 @ 12\" OC for multi-story.",
            "type": "choice",
            "choices": [
                {"value": "2x6_12", "label": "2×6 @ 12\" OC"},
                {"value": "2x6_16", "label": "2×6 @ 16\" OC"},
                {"value": "2x4_16", "label": "2×4 @ 16\" OC"},
                {"value": "2x4_12", "label": "2×4 @ 12\" OC"},
            ]
        })
        questions.append({
            "id": "stud_spec_interior",
            "question": "What stud size/spacing for interior load-bearing walls?",
            "hint": "Check structural general notes. Often 2x4 @ 16\" OC or 2x6 @ 16\" OC.",
            "type": "choice",
            "choices": [
                {"value": "2x4_16", "label": "2×4 @ 16\" OC"},
                {"value": "2x4_12", "label": "2×4 @ 12\" OC"},
                {"value": "2x6_16", "label": "2×6 @ 16\" OC"},
            ]
        })
        log("warn", "❓", "Stud specs not found — will ask user")

    # Ask about floor-to-floor height if not in truss notes
    if not session["floor_heights"]:
        questions.append({
            "id": "ftf_height",
            "question": "What is the typical floor-to-floor height? (Exterior walls)",
            "hint": "Check typical wall section or building section drawings.",
            "type": "choice",
            "choices": [
                {"value": "9",  "label": "9'-0\""},
                {"value": "10", "label": "10'-0\""},
                {"value": "12", "label": "12'-0\""},
                {"value": "other", "label": "Other (I'll type it)"},
            ]
        })
        log("warn", "❓", "Floor-to-floor height not found — will ask user")

    # Number of stories (to know load stack for stud sizing)
    n_levels = len(set(
        s["sheet_id"][:2] for s in session["sheets"]
        if s["sheet_id"] and re.match(r"[AS]\d", s["sheet_id"])
    ))
    if n_levels == 0:
        questions.append({
            "id": "num_stories",
            "question": "How many stories is this building?",
            "type": "choice",
            "choices": [
                {"value": "1", "label": "1 story"},
                {"value": "2", "label": "2 stories"},
                {"value": "3", "label": "3 stories"},
                {"value": "4", "label": "4 stories"},
                {"value": "5", "label": "5 stories"},
            ]
        })

    session["questions"] = questions
    return session


def apply_qa_answers(session: dict, answers: dict) -> dict:
    """Merge user Q&A answers back into the session."""
    # Truss depth
    if "truss_depth" in answers:
        session["truss_info"]["depth_in"] = int(answers["truss_depth"])

    # Stud specs
    for key in ("stud_spec_exterior", "stud_spec_interior"):
        if key in answers:
            wall_type = "exterior" if "exterior" in key else "interior"
            val = answers[key]  # e.g. "2x6_12"
            parts = val.split("_")
            if len(parts) == 2:
                size, oc = parts[0], int(parts[1])
                if wall_type not in session["stud_specs"]:
                    session["stud_specs"][wall_type] = {}
                # Apply to all load levels
                for load in range(5):
                    session["stud_specs"][wall_type][load] = (size, oc, 1)

    # Floor-to-floor height
    if "ftf_height" in answers:
        val = answers["ftf_height"]
        if val != "other":
            ft = int(val)
            for lvl in ["L1", "L2", "L3", "L4", "L5"]:
                session["floor_heights"][lvl] = ft * 12
    if "ftf_height_custom" in answers:
        try:
            in_val = parse_ft_in(answers["ftf_height_custom"]) or float(answers["ftf_height_custom"]) * 12
            for lvl in ["L1", "L2", "L3", "L4", "L5"]:
                session["floor_heights"][lvl] = in_val
        except Exception:
            pass

    return session


# ─────────────────────────────────────────────────────────────────────────────
# Schedule parsers (called after vision extraction)
# ─────────────────────────────────────────────────────────────────────────────

_PLY_RX      = re.compile(r"\((\d)\)")
_LUMBER_2X   = re.compile(r"\(?(\d)\)?\s*(\d{1,2})[xX](\d{1,2})")
_LUMBER_EWP  = re.compile(r'(\d+(?:-\d+/\d+)?)"?\s*[xX]\s*(\d+(?:-\d+/\d+)?)"?')
_MATERIAL_RX = re.compile(r"\b(PSL|LVL|LSL|GLULAM|SYP|SPF|DF|PT|FRT)\b", re.I)

def _stock_len_from_mark(mark: str) -> int:
    """Infer stock length from header mark like H4-12 → 12 ft."""
    m = re.search(r"-(\d{1,2})$", mark)
    return int(m.group(1)) if m else 10

def _parse_header_schedule(schedule: dict, out: dict):
    """Parse a vision-extracted header schedule into the out dict."""
    cols = schedule.get("columns", [])
    mark_col = next((c for c in cols if re.search(r"mark|tag|id", c, re.I)), None)
    size_col = next((c for c in cols if re.search(r"size|dimension|header", c, re.I)), None)
    if not mark_col:
        return
    for row in schedule.get("rows", []):
        mark = (row.get(mark_col) or "").strip()
        if not mark:
            continue
        raw_size = (row.get(size_col) or "") if size_col else ""
        # Parse size
        mat_m = _MATERIAL_RX.search(raw_size)
        is_ewp = bool(mat_m and mat_m.group(1).upper() in ("PSL", "LVL", "LSL", "GLULAM"))
        if is_ewp:
            m2 = _LUMBER_EWP.search(raw_size)
            lumber_size = f'{m2.group(1)}"x{m2.group(2)}"' if m2 else raw_size
            ply = int(_PLY_RX.search(raw_size).group(1)) if _PLY_RX.search(raw_size) else 1
        else:
            m2 = _LUMBER_2X.search(raw_size)
            ply = int(m2.group(1)) if m2 else 1
            lumber_size = f"{m2.group(2)}x{m2.group(3)}" if m2 else raw_size
        material = mat_m.group(1).upper() if mat_m else "SYP#2"
        stock_ft = _stock_len_from_mark(mark)
        out[mark] = {
            "lumber_size": lumber_size, "material": material,
            "ply_count": ply, "stock_ft": stock_ft,
            "flitch_plate": None,
        }

def _parse_shear_schedule(schedule: dict, out: dict):
    """Parse a shear wall schedule."""
    cols = schedule.get("columns", [])
    mark_col  = next((c for c in cols if re.search(r"mark|type", c, re.I)), None)
    sheath_col= next((c for c in cols if re.search(r"sheath|panel", c, re.I)), None)
    nail_col  = next((c for c in cols if re.search(r"nail|fast", c, re.I)), None)
    for i, row in enumerate(schedule.get("rows", [])):
        mark = (row.get(mark_col) or f"SW{i+1}").strip()
        out[mark] = {
            "sheathing": row.get(sheath_col),
            "nail":      row.get(nail_col),
        }


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _scale_label(ft_per_pt: float) -> str:
    """Convert ft/pt back to a readable scale string."""
    for name, val in {
        '1/4"=1\'': 1/18, '3/16"=1\'': 1/13.5,
        '1/8"=1\'': 1/9,  '1/2"=1\'': 1/36,
        '1"=1\'':   1/72,
    }.items():
        if abs(ft_per_pt - val) < 0.001:
            return name
    return f"{ft_per_pt:.5f}"
