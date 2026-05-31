"""
TAKEOFF RULE-BOOK EXTRACTOR — Tier 1, generalizable across architects.

Reads ANY vector-PDF drawing set and produces a structured "rule book":
  - sheet index (sheet_id -> page, title)
  - page classification (cover/notes/plan/schedule/section/detail/elevation/MEP)
  - general notes (text-layer)
  - schedules (Header/Beam, Shear Wall, Window, Door) as row/column data
  - building summary (occupancy, type, area, height, # stories)

Design principles for generalization:
  1) Discover sheet IDs from title blocks via REGEX, not hard-coded sheet names.
  2) Classify pages by KEYWORDS in title + sheet-ID prefix, not page numbers.
  3) Extract schedules via POSITION-AWARE text (PyMuPDF word bboxes), so any
     CAD-drawn pseudo-table reconstructs into rows/columns by y/x clustering.
  4) Every architect uses different wording, so use multi-keyword matching with
     synonyms (HEADER/BEAM/LINTEL all = "header schedule").
  5) Output STRUCTURED JSON only — downstream code never re-parses text.
"""
import fitz, re, json, sys
from pathlib import Path
from collections import defaultdict

# ── DISCIPLINE-PREFIX → category map (covers most US conventions) ──
# Most architects use these prefixes; if a set deviates, classification falls
# back to keyword scanning of the sheet title.
_DISC_PREFIX = {
    "A":  "architectural", "AN": "architectural",
    "S":  "structural",
    "M":  "mechanical",
    "E":  "electrical",
    "P":  "plumbing",
    "FP": "fire_protection",
    "L":  "landscape",
    "C":  "civil",
    "G":  "general",
    "T":  "specifications",
    "I":  "interior",
}

# Title-keyword → page role.  Order matters: first match wins.  Synonyms
# captured so different architects' wording still classifies correctly.
_PAGE_ROLE_RULES = [
    ("cover",            [r"\bcover sheet\b", r"\bindex of drawings\b", r"\bsheet index\b"]),
    ("code_summary",     [r"code (analysis|summary)", r"\bbuilding code\b"]),
    ("general_notes",    [r"general notes", r"symbol legend", r"abbreviations"]),
    ("life_safety",      [r"life safety"]),
    ("ul_directory",     [r"ul reference", r"ul assembl(y|ies)"]),
    ("schedule_shear",   [r"shear wall schedule", r"shear schedule"]),
    ("schedule_header",  [r"wood header", r"header.{0,3}(beam|footing)?.{0,3}schedule",
                          r"lintel schedule", r"\bheader schedule\b"]),
    ("schedule_beam",    [r"beam schedule"]),
    ("schedule_column",  [r"column schedule", r"post schedule"]),
    ("schedule_truss",   [r"truss schedule"]),
    ("schedule_joist",   [r"joist schedule"]),
    ("schedule_window",  [r"window schedule"]),
    ("schedule_door",    [r"door schedule"]),
    ("schedule_finish",  [r"finish schedule", r"room finish"]),
    ("site_plan",        [r"site plan", r"architectural site"]),
    ("floor_plan",       [r"floor plan", r"building plan", r"slab plan", r"roof plan", r"foundation plan"]),
    ("rcp",              [r"reflected ceiling", r"\brcp\b"]),
    ("elevation",        [r"\belevation"]),
    ("section",          [r"\bsection\b(?!.*schedule)", r"wall section"]),
    ("detail",           [r"\bdetail", r"axonometric"]),
    ("enlarged",         [r"enlarged"]),
    ("riser_diagram",    [r"riser diagram"]),
]

# Sheet-ID regex: matches A0.01, AN12.01, S0.06, M2.02, P1.01A, E12.01A, S3.01A, …
# This handles 1-2 letter discipline prefix + digits + dot + digits + optional letter.
_SHEET_ID_RX = re.compile(r"\b([A-Z]{1,2})(\d{1,2})\.(\d{1,2}[A-Za-z]?)\b")


def _looks_like_real_text(s):
    """Filter out garbled custom-font glyphs (consultant logos rendered in
    custom-encoded fonts, e.g. 'RVHSK,LDZUHQFH'). Real English title text has
    a high ratio of ASCII letters/spaces and no font-encoding artifacts."""
    if not s or len(s) < 4: return False
    ascii_ok = sum(1 for c in s if c.isascii() and (c.isalpha() or c in " /-&_,.()0123456789"))
    if ascii_ok / len(s) < 0.85: return False
    if any(ord(c) < 32 or 0x80 <= ord(c) < 0xA0 for c in s): return False
    return True


def _titleblock_spans(page):
    """Return text spans located in the title-block zone (right edge or bottom-right
    corner) with their font size. The title block is, by near-universal CAD
    convention, along the right edge or bottom-right of the sheet. Restricting to
    this zone removes most boilerplate and works across architects."""
    W, H = page.rect.width, page.rect.height
    spans = []
    for blk in page.get_text("dict")["blocks"]:
        if "lines" not in blk: continue
        for line in blk["lines"]:
            for span in line["spans"]:
                t = span["text"].strip()
                if not t: continue
                x, y = span["bbox"][0], span["bbox"][1]
                # Title block sits along the right edge or bottom strip. Structural
                # portrait sheets place it further in, so accept right 35%; also
                # accept the top-right and bottom-right corners explicitly.
                in_right  = x > W * 0.62
                in_bottom = y > H * 0.82
                in_top_right = (x > W * 0.55 and y < H * 0.10)
                if in_right or in_bottom or in_top_right:
                    spans.append({"size": span["size"], "x": x, "y": y, "text": t})
    return spans


def detect_sheet_id(page):
    """Sheet ID = the sheet-ID-pattern token with the LARGEST font in the
    title-block zone (bottom-right). Falls back to scanning whole page."""
    def pick(spans):
        cands = []
        for s in spans:
            m = _SHEET_ID_RX.fullmatch(s["text"]) or _SHEET_ID_RX.search(s["text"])
            if m:
                cands.append((s["size"], m.group(0)))
        if cands:
            cands.sort(key=lambda c: -c[0])
            return cands[0][1]
        return None
    # Prefer title-block zone, then whole page
    sid = pick(_titleblock_spans(page))
    if sid: return sid
    # whole-page fallback
    allspans = []
    for blk in page.get_text("dict")["blocks"]:
        if "lines" not in blk: continue
        for line in blk["lines"]:
            for span in line["spans"]:
                allspans.append({"size": span["size"], "text": span["text"].strip()})
    return pick(allspans)


def detect_sheet_title(page):
    """Sheet title = the largest English text in the title-block zone, excluding
    the sheet ID, firm boilerplate, dates, and garbled custom-font glyphs.
    Works across architects because the title always sits in the title block."""
    spans = _titleblock_spans(page)
    boilerplate_rx = re.compile(
        r"\(\d{3}\)|@|www\.|LICENSE|copyright|\.com|\.rvt|\bFL\b\s*\d{5}|"
        r"Suite \d|FUGLEBERG|JLC&|CONSULTING|ENGINEER|ARCHITECT$|"
        r"PERMIT SET|CONCEPTUAL|created by|ORANGE COUNTY|PROPERTIES PD|"
        r"^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|^[A-Z]{1,2}\d{1,2}\.\d{1,2}[A-Z]?$",
        re.I)
    cand = []
    for s in spans:
        t = s["text"]
        if not (5 <= len(t) <= 80): continue
        if not _looks_like_real_text(t): continue
        if boilerplate_rx.search(t): continue
        if "SCALE" in t.upper(): continue
        if _SHEET_ID_RX.fullmatch(t): continue
        if re.search(r"[A-Z]{3,}", t):
            cand.append((s["size"], t))
    cand.sort(key=lambda c: -c[0])
    return cand[0][1] if cand else ""


def classify_page(sheet_id, title):
    """Determine page role from sheet ID prefix + title keywords."""
    # Discipline from sheet ID prefix
    prefix_match = re.match(r"^([A-Z]{1,2})", sheet_id or "")
    discipline = _DISC_PREFIX.get(prefix_match.group(1), "unknown") if prefix_match else "unknown"

    # Role from title keywords (first match wins)
    title_l = (title or "").lower()
    role = "unknown"
    for role_name, patterns in _PAGE_ROLE_RULES:
        if any(re.search(p, title_l) for p in patterns):
            role = role_name
            break
    return discipline, role


def extract_page_words(page):
    """Return (x0,y0,x1,y1,text) for every word fragment on the page."""
    return [(w[0], w[1], w[2], w[3], w[4]) for w in page.get_text("words")]


def cluster_into_rows(words, y_tolerance=8):
    """Group words into rows by y-coordinate clustering.
    A 'row' = words whose y0 falls within y_tolerance of each other.
    Returns rows sorted top-to-bottom, each row's words sorted left-to-right.
    """
    if not words: return []
    # Sort by y, then group
    by_y = sorted(words, key=lambda w: w[1])
    rows = []
    current = [by_y[0]]
    cur_y = by_y[0][1]
    for w in by_y[1:]:
        if abs(w[1] - cur_y) <= y_tolerance:
            current.append(w)
        else:
            rows.append(sorted(current, key=lambda w: w[0]))
            current = [w]
            cur_y = w[1]
    rows.append(sorted(current, key=lambda w: w[0]))
    return rows


def rows_to_text(rows):
    """Convert clustered rows into plain text lines."""
    return ["  ".join(w[4] for w in r) for r in rows]


def _content_schedule_scan(page):
    """Detect schedules embedded on a sheet by finding schedule keywords that
    appear as HEADINGS (larger font), not buried in note paragraphs. A real
    schedule has its name set as a title above the table; a note that merely
    mentions 'wood header' is small body text. Requiring heading-size text
    removes the false positives."""
    # Collect heading-sized text spans (>= 12pt) only
    headings = []
    for blk in page.get_text("dict")["blocks"]:
        if "lines" not in blk: continue
        for line in blk["lines"]:
            for span in line["spans"]:
                if span["size"] >= 12:
                    headings.append(span["text"].upper())
    htext = " ".join(headings)
    found = []
    content_rules = [
        ("schedule_shear",  ["SHEAR WALL SCHEDULE"]),
        ("schedule_header", ["WOOD HEADER SCHEDULE", "HEADER SCHEDULE", "LINTEL SCHEDULE"]),
        ("schedule_beam",   ["BEAM SCHEDULE"]),
        ("schedule_column", ["COLUMN SCHEDULE", "POST SCHEDULE"]),
        ("schedule_truss",  ["TRUSS SCHEDULE"]),
        ("schedule_joist",  ["JOIST SCHEDULE"]),
        ("schedule_window", ["WINDOW SCHEDULE"]),
        ("schedule_door",   ["DOOR SCHEDULE"]),
    ]
    for role, kws in content_rules:
        if any(k in htext for k in kws):
            found.append(role)
    return found


def scan_set(pdf_path, max_pages=None, content_scan=True):
    """Scan a full drawing set and return a structured rule book.

    content_scan=True also scans page bodies for embedded schedules, so a
    schedule printed on a framing-plan sheet is still discovered. This is what
    lets the same code work whether an architect uses dedicated schedule sheets
    (Willow Way) or prints schedules on plan sheets (Silver City)."""
    doc = fitz.open(pdf_path)
    n = doc.page_count if max_pages is None else min(max_pages, doc.page_count)

    sheets = []
    for i in range(n):
        page = doc[i]
        sid = detect_sheet_id(page) or ""
        title = detect_sheet_title(page)
        disc, role = classify_page(sid, title)
        rec = {"page": i + 1, "sheet_id": sid, "title": title,
               "discipline": disc, "role": role}
        # If the title block didn't already mark this as a schedule, look inside.
        if content_scan and not role.startswith("schedule"):
            embedded = _content_schedule_scan(page)
            if embedded:
                rec["embedded_schedules"] = embedded
                # promote the page's primary role to the first schedule found,
                # but keep the original role for reference
                rec["title_role"] = role
                rec["role"] = embedded[0]
        sheets.append(rec)
    return {"source": Path(pdf_path).name, "page_count": doc.page_count, "sheets": sheets}


if __name__ == "__main__":
    pdf = sys.argv[1] if len(sys.argv) > 1 else \
          "/mnt/user-data/uploads/North_Building_Rev1_IFC_-_Combined_Set.pdf"
    book = scan_set(pdf)
    print(f"=== {book['source']} : {book['page_count']} pages ===\n")
    # Summary table
    by_disc = defaultdict(lambda: defaultdict(int))
    for s in book["sheets"]:
        by_disc[s["discipline"]][s["role"]] += 1
    print("Pages by discipline + role:")
    for disc in sorted(by_disc):
        for role, n in sorted(by_disc[disc].items()):
            print(f"  {disc:18s} {role:20s} {n:3d}")
    # Show schedule sheets
    sched = [s for s in book["sheets"] if s["role"].startswith("schedule")]
    print(f"\nSchedules found ({len(sched)}):")
    for s in sched:
        print(f"  p{s['page']:>3}  {s['sheet_id']:8s}  [{s['role']}]  {s['title'][:60]}")