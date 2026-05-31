"""
VISION-BASED SCHEDULE READER — extracts schedule contents from any architect's drawings.

Why vision instead of CAD text reconstruction:
  CAD-drawn schedules look like tables but their text is positioned freely in
  (x,y) space — there is no real grid. Architect A may put 4 schedules
  side-by-side on one sheet (Willow Way S0.06); architect B may put one
  schedule centered; architect C may use horizontal vs vertical orientation.
  Trying to reverse-engineer the grid is brittle. A vision model reads the
  schedule the same way a human estimator does, regardless of layout.

Generalization principle:
  The PROMPT never assumes specific marks, sizes, or materials. It asks the
  model to extract "whatever rows exist in whatever schedule(s) you see" and
  preserve the architect's actual values verbatim. This makes the reader work
  on Willow Way (FK Architecture), on the Silver City project, or on any
  future job — same code, different architect.

Public API:
    read_schedule_page(pdf_path, page_number, dpi=200) -> dict
    read_all_schedules(pdf_path, scan_result) -> dict
"""
import fitz, base64, json, re, os, subprocess, tempfile
from pathlib import Path

# ── Prompt template (architect-agnostic) ─────────────────────────────────
# Designed so it works whether the schedule contains H2-10 marks, SW-1
# marks, custom mark naming, vertical or horizontal layouts.
_SCHEDULE_PROMPT = """You are reading a construction drawing sheet that contains one or more SCHEDULES (tables of building elements with their specs). This sheet may contain multiple schedules side-by-side — read every schedule on the page.

For each schedule on this page, return a JSON object with:
  - "schedule_name": the exact title of the schedule as written (e.g. "HEADER/BEAM SCHEDULE", "SHEAR WALL SCHEDULE", "WINDOW SCHEDULE", "DOOR SCHEDULE", "FOOTING SCHEDULE"). Use the EXACT wording from the drawing.
  - "columns": an array of the exact column header names (e.g. ["MARK", "HEADER SIZE", "FLITCH PLATE", "NO. OF PLIES", "ASSEMBLY PATTERN"]). Use the EXACT wording.
  - "rows": an array of objects, one per data row. Each row object maps the column name (verbatim) to the cell value (verbatim). Preserve "1/2" PLY", "5-1/4" PSL", "3-1/2"x11-7/8"", etc. exactly as written including quotes and slashes.

Important rules:
  - Do NOT assume which schedules will be on the page. Read whatever is actually there.
  - Do NOT normalize values — keep the architect's original wording (e.g. if it says "(2) 2x8" leave it that way).
  - If a cell has a dash "-" or is empty, return null for that cell.
  - Ignore figures/details/elevations on the same sheet that are not schedules (skip these — only return actual tabular schedules).
  - If you cannot read a cell with confidence, return the cell as null and add a notes entry.

Return your output as a JSON object with this shape:
{
  "schedules": [
    { "schedule_name": "...", "columns": [...], "rows": [{...}, {...}] }
  ],
  "notes": ["any general notes you noticed printed on the schedule, e.g. 'PROVIDE WOOD HEADERS OVER ALL OPENINGS'"]
}

Respond with ONLY the JSON object, no other text."""


def render_page_jpeg(pdf_path, page_number, dpi=200, out_dir=None):
    """Render one page to a JPEG file and return the path."""
    out_dir = Path(out_dir or tempfile.gettempdir())
    out_dir.mkdir(parents=True, exist_ok=True)
    prefix = out_dir / f"sched_p{page_number:03d}"
    subprocess.run(
        ["pdftoppm", "-jpeg", "-r", str(dpi),
         "-f", str(page_number), "-l", str(page_number),
         str(pdf_path), str(prefix)],
        check=True, capture_output=True)
    matches = sorted(out_dir.glob(f"sched_p{page_number:03d}-*.jpg"))
    return matches[0] if matches else None


def read_schedule_page(pdf_path, page_number, dpi=200, client=None):
    """Render page, send to Claude vision, parse and return structured schedules.

    `client`: optional pre-built anthropic.Anthropic() client. If None, will
              attempt to create one (requires ANTHROPIC_API_KEY in env).

    Returns: {"page": int, "schedules": [...], "notes": [...], "error": ...}
    """
    img_path = render_page_jpeg(pdf_path, page_number, dpi=dpi)
    if not img_path:
        return {"page": page_number, "error": "render_failed", "schedules": []}

    with open(img_path, "rb") as f:
        img_b64 = base64.standard_b64encode(f.read()).decode()

    # Try to call the API. If no key / no network, return the prepared payload
    # so the caller can run it on their server with credentials.
    try:
        import anthropic
        client = client or anthropic.Anthropic()
        resp = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=8000,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {
                        "type": "base64", "media_type": "image/jpeg", "data": img_b64}},
                    {"type": "text", "text": _SCHEDULE_PROMPT},
                ],
            }],
        )
        text = "".join(b.text for b in resp.content if b.type == "text")
        # Tolerate ```json fences
        clean = re.sub(r"^```(?:json)?\s*|\s*```\s*$", "", text.strip(), flags=re.M)
        data = json.loads(clean)
        return {
            "page": page_number,
            "schedules": data.get("schedules", []),
            "notes":     data.get("notes", []),
            "_image":    str(img_path),
        }
    except Exception as e:
        return {
            "page": page_number,
            "error": str(e),
            "_image": str(img_path),
            "schedules": [],
        }


def read_all_schedules(pdf_path, scan_result, dpi=200, client=None):
    """Read every schedule-classified page in a scanned set.

    `scan_result`: output from extractor.scan_set(pdf_path)
    Returns a list of page-results, each with its schedules and source info.
    """
    out = []
    for s in scan_result["sheets"]:
        if not s["role"].startswith("schedule"): continue
        result = read_schedule_page(pdf_path, s["page"], dpi=dpi, client=client)
        result["sheet_id"]    = s["sheet_id"]
        result["sheet_title"] = s["title"]
        result["page_role"]   = s["role"]
        out.append(result)
    return out


# ────────────────────────────────────────────────────────────────────────
#                       OFFLINE VERIFICATION HARNESS
# ────────────────────────────────────────────────────────────────────────
# When ANTHROPIC_API_KEY isn't available (e.g. testing locally), simulate the
# vision response from a known-good ground-truth dict. This lets us verify
# the full pipeline structure (render -> call -> parse -> downstream rule
# book) without hitting the API. In production this branch is never taken.

_GROUND_TRUTH_FALLBACK = {
    # Willow Way S0.06 - HEADER/BEAM SCHEDULE (page 93 of North Building)
    ("North_Building_Rev1_IFC_-_Combined_Set.pdf", 93): {
        "schedules": [
            {
                "schedule_name": "HEADER/BEAM SCHEDULE",
                "columns": ["MARK", "HEADER SIZE", "FLITCH PLATE",
                            "NO. OF PLIES", "ASSEMBLY PATTERN"],
                "rows": [
                    {"MARK": "H2-08", "HEADER SIZE": "(2) 2x8",  "FLITCH PLATE": "1/2\" PLY", "NO. OF PLIES": "3", "ASSEMBLY PATTERN": "A,B"},
                    {"MARK": "H2-10", "HEADER SIZE": "(2) 2x10", "FLITCH PLATE": "1/2\" PLY", "NO. OF PLIES": "3", "ASSEMBLY PATTERN": "A,B"},
                    {"MARK": "H2-12", "HEADER SIZE": "(2) 2x12", "FLITCH PLATE": "1/2\" PLY", "NO. OF PLIES": "3", "ASSEMBLY PATTERN": "A,B"},
                    {"MARK": "H3-08", "HEADER SIZE": "(3) 2x8",  "FLITCH PLATE": "1/2\" PLY", "NO. OF PLIES": "3", "ASSEMBLY PATTERN": "F,G"},
                    {"MARK": "H3-10", "HEADER SIZE": "(3) 2x10", "FLITCH PLATE": "1/2\" PLY", "NO. OF PLIES": "3", "ASSEMBLY PATTERN": "F,G"},
                    {"MARK": "H3-12", "HEADER SIZE": "(3) 2x12", "FLITCH PLATE": "1/2\" PLY", "NO. OF PLIES": "3", "ASSEMBLY PATTERN": "F,G"},
                    {"MARK": "H4-08", "HEADER SIZE": "3-1/2\"x7-1/4\" PSL",  "FLITCH PLATE": "-", "NO. OF PLIES": "-", "ASSEMBLY PATTERN": "A,B"},
                    {"MARK": "H4-10", "HEADER SIZE": "3-1/2\"x9-1/4\" PSL",  "FLITCH PLATE": "-", "NO. OF PLIES": "-", "ASSEMBLY PATTERN": "A,B"},
                    {"MARK": "H4-12", "HEADER SIZE": "3-1/2\"x11-7/8\" PSL", "FLITCH PLATE": "-", "NO. OF PLIES": "-", "ASSEMBLY PATTERN": "C,D"},
                    {"MARK": "H4-14", "HEADER SIZE": "3-1/2\"x14\" PSL",     "FLITCH PLATE": "-", "NO. OF PLIES": "-", "ASSEMBLY PATTERN": "C,D"},
                    {"MARK": "H4-16", "HEADER SIZE": "3-1/2\"x16\" PSL",     "FLITCH PLATE": "-", "NO. OF PLIES": "-", "ASSEMBLY PATTERN": "C,D"},
                    {"MARK": "H4-18", "HEADER SIZE": "3-1/2\"x18\" PSL",     "FLITCH PLATE": "-", "NO. OF PLIES": "-", "ASSEMBLY PATTERN": "C,D"},
                    {"MARK": "H4-20", "HEADER SIZE": "(3) 1-3/4\"x20\" LVL", "FLITCH PLATE": "-", "NO. OF PLIES": "-", "ASSEMBLY PATTERN": "E"},
                    {"MARK": "H6-10", "HEADER SIZE": "5-1/4\"x7-1/4\" PSL",  "FLITCH PLATE": "-", "NO. OF PLIES": "-", "ASSEMBLY PATTERN": "F,G"},
                    {"MARK": "H6-12", "HEADER SIZE": "5-1/4\"x9-1/4\" PSL",  "FLITCH PLATE": "-", "NO. OF PLIES": "-", "ASSEMBLY PATTERN": "H,J"},
                    {"MARK": "H6-14", "HEADER SIZE": "5-1/4\"x11-7/8\" PSL", "FLITCH PLATE": "-", "NO. OF PLIES": "-", "ASSEMBLY PATTERN": "H,J"},
                    {"MARK": "H6-16", "HEADER SIZE": "5-1/4\"x14\" PSL",     "FLITCH PLATE": "-", "NO. OF PLIES": "-", "ASSEMBLY PATTERN": "H,J"},
                    {"MARK": "H6-18", "HEADER SIZE": "5-1/4\"x16\" PSL",     "FLITCH PLATE": "-", "NO. OF PLIES": "-", "ASSEMBLY PATTERN": "H,J"},
                    {"MARK": "H6-20", "HEADER SIZE": "(3) 1-3/4\"x20\" LVL", "FLITCH PLATE": "-", "NO. OF PLIES": "-", "ASSEMBLY PATTERN": "K,L"},
                    {"MARK": "H8-10", "HEADER SIZE": "7\"x9-1/4\" PSL",  "FLITCH PLATE": "-", "NO. OF PLIES": "-", "ASSEMBLY PATTERN": "K,L"},
                    {"MARK": "H8-12", "HEADER SIZE": "7\"x11-7/8\" PSL", "FLITCH PLATE": "-", "NO. OF PLIES": "-", "ASSEMBLY PATTERN": "K,L"},
                    {"MARK": "H8-14", "HEADER SIZE": "7\"x14\" PSL", "FLITCH PLATE": "-", "NO. OF PLIES": "-", "ASSEMBLY PATTERN": "K,L"},
                    {"MARK": "H8-16", "HEADER SIZE": "7\"x16\" PSL", "FLITCH PLATE": "-", "NO. OF PLIES": "-", "ASSEMBLY PATTERN": "K,L"},
                    {"MARK": "H8-18", "HEADER SIZE": "7\"x18\" PSL", "FLITCH PLATE": "-", "NO. OF PLIES": "-", "ASSEMBLY PATTERN": "K,L"},
                    {"MARK": "H8-20", "HEADER SIZE": "(4) 1-3/4\"x20\" LVL", "FLITCH PLATE": "-", "NO. OF PLIES": "-", "ASSEMBLY PATTERN": "M"},
                ],
            },
            {
                "schedule_name": "FOOTING SCHEDULE",
                "columns": ["MARK", "SIZE WIDTH/LENGTH/DEPTH",
                            "REINFORCING BOTTOM", "REINFORCING TOP", "REMARKS"],
                "rows": [
                    {"MARK": "F-2.0",   "SIZE WIDTH/LENGTH/DEPTH": "2'-0\"x2'-0\"x1'-0\"",  "REINFORCING BOTTOM": "(3) #5 EA. WAY", "REINFORCING TOP": None, "REMARKS": None},
                    {"MARK": "F-2.5",   "SIZE WIDTH/LENGTH/DEPTH": "2'-6\"x2'-6\"x1'-0\"",  "REINFORCING BOTTOM": "(3) #5 EA. WAY", "REINFORCING TOP": None, "REMARKS": None},
                    {"MARK": "F-3.0",   "SIZE WIDTH/LENGTH/DEPTH": "3'-0\"x3'-0\"x1'-2\"",  "REINFORCING BOTTOM": "(4) #5 EA. WAY", "REINFORCING TOP": None, "REMARKS": None},
                    {"MARK": "F-4.0",   "SIZE WIDTH/LENGTH/DEPTH": "4'-0\"x4'-0\"x1'-2\"",  "REINFORCING BOTTOM": "(5) #5 EA. WAY", "REINFORCING TOP": None, "REMARKS": None},
                    {"MARK": "F-5.0",   "SIZE WIDTH/LENGTH/DEPTH": "5'-0\"x5'-0\"x1'-2\"",  "REINFORCING BOTTOM": "(6) #5 EA. WAY", "REINFORCING TOP": None, "REMARKS": None},
                    {"MARK": "F-6.0x4.0","SIZE WIDTH/LENGTH/DEPTH": "6'-0\"x4'-0\"x1'-4\"", "REINFORCING BOTTOM": "#5 AT 12\" OC EA. WAY", "REINFORCING TOP": None, "REMARKS": None},
                    {"MARK": "WF-2.0",  "SIZE WIDTH/LENGTH/DEPTH": "3'-0\"xCONT.x1'-2\"",   "REINFORCING BOTTOM": "(4) #5 CONT.", "REINFORCING TOP": "#3 STIRRUPS AT 24\" OC", "REMARKS": None},
                ],
            },
        ],
        "notes": [
            "PROVIDE WOOD HEADERS OVER ALL OPENINGS. IF NO HEADER IS SPECIFIED, PROVIDE H2-10 AT EXTERIOR WALLS AND PROVIDE H2-12 AT ALL INTERIOR LOAD-BEARING WALLS.",
            "FLITCH PLATES SHALL BE PROVIDED BETWEEN HEADER PLIES.",
            "ASSEMBLY PATTERN PER ENGINEER WOOD LUMBER APPLIES ONLY IF THE BEAM HAS MULTIPLE PLIES.",
        ],
    },
}


def read_schedule_page_offline(pdf_path, page_number, dpi=200):
    """Same return shape as read_schedule_page, but uses ground-truth fallback
    when no API key is available. For DEV/TEST only — in production
    read_schedule_page() makes the real API call."""
    key = (Path(pdf_path).name, page_number)
    truth = _GROUND_TRUTH_FALLBACK.get(key)
    img_path = render_page_jpeg(pdf_path, page_number, dpi=dpi)
    if truth:
        return {"page": page_number, "_image": str(img_path) if img_path else None,
                "_offline": True, **truth}
    return {"page": page_number, "_image": str(img_path) if img_path else None,
            "_offline": True, "schedules": [],
            "notes": ["(offline mode: no ground-truth available for this page)"]}


if __name__ == "__main__":
    # Self-test: render and (try to) read the S0.06 page from North Building.
    import sys
    sys.path.insert(0, str(Path(__file__).parent))
    from extractor import scan_set

    pdf = "/mnt/user-data/uploads/North_Building_Rev1_IFC_-_Combined_Set.pdf"
    scan = scan_set(pdf)
    schedules = [s for s in scan["sheets"] if s["role"].startswith("schedule")]
    print(f"{Path(pdf).name}: {len(schedules)} schedule pages")
    for s in schedules:
        print(f"  p{s['page']} {s['sheet_id']} [{s['role']}]")

    # Try real API call first; fall back to offline-truth on auth/network error
    use_offline = not os.environ.get("ANTHROPIC_API_KEY")
    if use_offline:
        print("\n(no ANTHROPIC_API_KEY — using offline ground-truth simulator)\n")
        target = next(s for s in schedules if s["sheet_id"] == "S0.06")
        out = read_schedule_page_offline(pdf, target["page"])
    else:
        target = next(s for s in schedules if s["sheet_id"] == "S0.06")
        out = read_schedule_page(pdf, target["page"])

    print(f"=== Page {out['page']} — schedules found: {len(out.get('schedules', []))} ===")
    for sch in out.get("schedules", []):
        print(f"\n● {sch['schedule_name']}")
        print(f"  Columns: {sch['columns']}")
        print(f"  Rows: {len(sch['rows'])}")
        for row in sch['rows'][:5]:
            print(f"    {row}")
        if len(sch['rows']) > 5:
            print(f"    ... +{len(sch['rows']) - 5} more rows")
    if out.get("notes"):
        print(f"\nGeneral notes captured: {len(out['notes'])}")
        for n in out["notes"][:3]:
            print(f"  - {n[:90]}")
