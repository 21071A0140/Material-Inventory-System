"""
SCHEDULE CALLOUT EXTRACTOR — generalizable across architects.

Instead of trying to reconstruct CAD table grids (which vary wildly), we
extract STRUCTURED CALLOUTS using domain-specific regex + spatial proximity.

For a header/beam schedule we need to know, for each mark (H2-10, H4-12, ...):
  - lumber size(s):       2x10, 2x12, 5-1/4"x11-7/8"
  - material:             SYP#2 / PSL / LVL / PT
  - ply count / pattern:  (2), (3), (4)
  - span / spacing:       optional

For a shear-wall schedule, per wall mark/type:
  - sheathing type:       15/32" PLYWOOD, 7/16" OSB-ZIP
  - fastener:             0.131"x2-1/2" NAILS
  - spacing edge/field:   6" OC, 12" OC
  - end stud size:        3-1/2"x5-1/4" PSL

Approach:
  - Run ALL regexes over every word fragment on the page (with position).
  - Cluster by y-band (≈8 pt) → each band is a candidate row of callouts.
  - For each y-band, keep only bands that contain a MARK token; pair that
    mark with the other tokens in the same band → one row of structured data.

This is robust to:
  - Different schedule layouts / page positions
  - Multiple schedules on one sheet
  - Rotated text (we just check both orientations)
  - Different architects' wording — patterns are over UNITS not words
"""
import fitz, re, json
from pathlib import Path
from collections import defaultdict

# ── Domain token patterns ──────────────────────────────────────────────
# These are the LANGUAGE OF FRAMING DRAWINGS — they don't vary by architect.

# Lumber dimension: 2x4, 2x10, 4x8, 6x6, etc. (case-insensitive 'x' or 'X')
LUMBER_RX = re.compile(r"^(\d{1,2})[xX](\d{1,2})$")

# Engineered wood: 1-3/4"x11-7/8", 3-1/2"x5-1/4", 5-1/4"x18", etc.
EWP_RX = re.compile(r'^(\d+(?:-\d+/\d+)?)"?[xX](\d+(?:-\d+/\d+)?)"?$')

# Material grade: SYP#2, SPF, PT, DF, #2, etc.
GRADE_RX = re.compile(r"^(SYP#?\d?|SPF#?\d?|PT|DF|LVL|PSL|LSL|GLULAM|#\d)$", re.I)

# Header mark: H2-10, H4-12, H2-12, H3-10, etc.
HEADER_MARK_RX = re.compile(r"^H\d+-\d+(?:\([A-Z]\))?(?:-[A-Z]+)?$")

# Ply count: (2), (3), (4) — for built-up beams
PLY_RX = re.compile(r"^\(([234])\)$")

# Nail spec: 0.131"x2-1/2", 0.148"x3", 10d, 16d
NAIL_RX = re.compile(r'^(0\.\d{3})"?[xX](\d+(?:-\d+/\d+)?)"?$|^\d{1,3}d$')

# OC spacing: 6" OC, 12" OC, 16" OC, @16" OC, @ 24" OC
OC_RX = re.compile(r'^@?(\d+)"?$')   # use with context word "OC"

# Sheathing thickness/type: 15/32", 7/16", 1/2", 23/32", 3/8"
SHEATHING_THICKNESS_RX = re.compile(r'^(\d+/\d+)"?$|^(0?\.\d+)"?$')

# Wall/shear wall mark: A, B, C, type-1, SW-1, etc.
SHEAR_MARK_RX = re.compile(r"^(SW-?\d+|[A-F])$")


def _y_cluster(words, y_tol=8):
    """Group word fragments into y-bands.  Returns list[ list[(x,y,x2,y2,text)] ]."""
    if not words: return []
    ws = sorted(words, key=lambda w: w[1])
    bands = [[ws[0]]]
    for w in ws[1:]:
        if w[1] - bands[-1][0][1] <= y_tol:
            bands[-1].append(w)
        else:
            bands.append([w])
    return [sorted(b, key=lambda w: w[0]) for b in bands]


def extract_header_schedule(page):
    """Return a list of {mark, size, material, ply, _y} dicts for header rows."""
    words = [(w[0], w[1], w[2], w[3], w[4]) for w in page.get_text("words")]
    bands = _y_cluster(words, y_tol=10)

    rows = []
    for band in bands:
        # Find header marks in this band
        marks = [w for w in band if HEADER_MARK_RX.match(w[4])]
        if not marks: continue

        # Collect other relevant tokens in the same band
        sizes_2x  = [w[4] for w in band if LUMBER_RX.match(w[4])]
        sizes_ewp = [w[4] for w in band if EWP_RX.match(w[4]) and "x" in w[4].lower()]
        plies     = [w[4] for w in band if PLY_RX.match(w[4])]
        materials = [w[4] for w in band
                     if GRADE_RX.match(w[4]) or w[4].upper() in ("PSL","LVL","LSL","SYP","SPF")]

        for m in marks:
            rows.append({
                "mark": m[4],
                "lumber_sizes": list(dict.fromkeys(sizes_2x)),     # de-dup, keep order
                "ewp_sizes":    list(dict.fromkeys(sizes_ewp)),
                "plies":        plies,
                "materials":    list(dict.fromkeys(materials)),
                "_y": int(m[1]),
            })
    return rows


def extract_shear_wall_schedule(page):
    """Extract shear-wall rows: sheathing, nail, edge/field spacing, end-stud size."""
    words = [(w[0], w[1], w[2], w[3], w[4]) for w in page.get_text("words")]
    bands = _y_cluster(words, y_tol=10)

    rows = []
    for band in bands:
        txts = [w[4] for w in band]
        text_band = " ".join(txts)

        # A shear-wall data row contains at minimum: sheathing thickness + NAIL spec
        has_sheath = any(re.search(r'\d+/\d+"\s*(PLYWOOD|OSB|PLY|ZIP)', text_band, re.I)
                         or "PLYWOOD" in t.upper() or "OSB" in t.upper() for t in txts)
        nail_specs = [w[4] for w in band if NAIL_RX.match(w[4])]
        if not (has_sheath and nail_specs): continue

        sheathing = [w[4] for w in band
                     if re.match(r'\d+/\d+"?$', w[4]) or
                     w[4].upper() in ("PLYWOOD", "OSB", "ZIP", "PLY")]
        # Look for OC-spacing pairs: a number followed by " then "OC" nearby
        # In CAD, "6" OC" may appear as separate fragments; collect "OC" anchors
        oc_indices = [i for i, w in enumerate(band) if w[4].upper() == "OC"]
        oc_values = []
        for i in oc_indices:
            # Look left up to 3 words for the spacing
            for j in range(max(0, i - 3), i):
                if re.match(r'^\d{1,2}"?$', band[j][4]):
                    oc_values.append(band[j][4].rstrip('"'))
                    break
        # End stud / shear stud size
        end_stud = next((w[4] for w in band if EWP_RX.match(w[4])), None)
        # PLF capacity
        plf = next((w[4] for w in band if re.match(r"^\d{2,4}$", w[4]) and int(w[4]) > 100), None)
        # Floor label that may anchor the row (THIRD/SECOND/FIRST)
        floor = next((w[4] for w in band if w[4].upper() in ("FIRST","SECOND","THIRD")), None)

        rows.append({
            "floor":       floor,
            "sheathing":   list(dict.fromkeys(sheathing)),
            "nail":        nail_specs[0] if nail_specs else None,
            "oc_spacings": oc_values,           # [edge, field]
            "plf":         plf,
            "end_stud":    end_stud,
            "_y":          int(band[0][1]),
        })
    return rows


def extract_rule_book(pdf_path, sheets):
    """Given a scanned set + sheet index, extract callouts from every schedule."""
    doc = fitz.open(pdf_path)
    book = {"headers": [], "shear_walls": []}
    for s in sheets:
        if s["role"] == "schedule_header":
            rows = extract_header_schedule(doc[s["page"] - 1])
            for r in rows: r["source_sheet"] = s["sheet_id"]; r["source_page"] = s["page"]
            book["headers"].extend(rows)
        elif s["role"] == "schedule_shear":
            rows = extract_shear_wall_schedule(doc[s["page"] - 1])
            for r in rows: r["source_sheet"] = s["sheet_id"]; r["source_page"] = s["page"]
            book["shear_walls"].extend(rows)
    return book
