"""
SWIFTJOB PARSER — read a PlanSwift .SwiftJob file and reconstruct the takeoff.

A .SwiftJob is a ZIP containing:
  - XMLData.XML : the full job tree (pages, scales, takeoff items, geometry)
  - NNN.tiff    : one raster image per page

This parser reads XMLData.XML and reconstructs, for every takeoff item:
  - the item name (which encodes scope + spec, e.g. "2x4 @ 16\" O.C - 16FT")
  - its type (Area / Linear / Count)
  - the page(s) it was drawn on + each page's scale
  - the measured quantity (LF / SF / count) recomputed from digitized points

Why this matters for the automation:
  This is GROUND TRUTH for how an estimator takes off a set. It tells us
  exactly which pages get measured (only ~37 of 294 here), what scale each
  uses, and how each measurement maps to a material. We use it to (a) learn
  the naming->material mapping and (b) validate the automated takeoff against
  a known-good human takeoff.

Generalization: the parser reads the PlanSwift schema, which is the same for
every job regardless of architect. Item NAMES vary, but the structure does not.
"""
import zipfile, math, re
import xml.etree.ElementTree as ET
from pathlib import Path


def _props(item):
    d = {}
    p = item.find("Properties")
    if p is not None:
        for pr in p.findall("Property"):
            d[pr.get("Name")] = (pr.text or "").strip()
    return d


def _parse_points(xmlstr):
    return [(float(x), float(y))
            for x, y in re.findall(r'X="([-\d.]+)"\s+Y="([-\d.]+)"', xmlstr or "")]


def _polyline_len(pts):
    return sum(math.dist(pts[i - 1], pts[i]) for i in range(1, len(pts)))


def _polygon_area(pts):
    # Shoelace; pts in pixel space
    if len(pts) < 3: return 0.0
    a = 0.0
    for i in range(len(pts)):
        x1, y1 = pts[i]; x2, y2 = pts[(i + 1) % len(pts)]
        a += x1 * y2 - x2 * y1
    return abs(a) / 2.0


def load_xml(swiftjob_path, workdir="/tmp/swiftjob"):
    z = zipfile.ZipFile(swiftjob_path)
    z.extract("XMLData.XML", workdir)
    return ET.parse(Path(workdir) / "XMLData.XML").getroot()


def build_page_index(root):
    """GUID -> {name, scale_x, scale_label, units}."""
    pages = {}
    def walk(item):
        pr = _props(item)
        if item.get("Class") == "Page" or pr.get("Type") == "Page":
            try: sx = float(pr.get("ScaleX") or 0)
            except Exception: sx = 0.0
            pages[pr.get("GUID", "")] = {
                "name": item.get("Name"),
                "scale_x": sx,
                "scale_label": pr.get("AutoScaled", ""),
                "units": pr.get("Scale Units", ""),
            }
        its = item.find("Items")
        if its is not None:
            for k in its.findall("Item"): walk(k)
    walk(root)
    return pages


def reconstruct_takeoff(swiftjob_path):
    """Return a list of takeoff records with measured quantities."""
    root = load_xml(swiftjob_path)
    pages = build_page_index(root)

    takeoff = None
    for it in root.find("Items").findall("Item"):
        if it.get("Name") == "Takeoff":
            takeoff = it; break
    if takeoff is None:
        return {"pages": pages, "items": []}

    records = []
    for it in takeoff.find("Items").findall("Item"):
        cls = it.get("Class")
        if cls not in ("Area", "Linear", "Count"):
            continue
        name = it.get("Name")
        kids = it.find("Items")
        secs = kids.findall("Item") if kids is not None else []

        total = 0.0
        page_names = set()
        scale_labels = set()
        for sec in secs:
            pr = _props(sec)
            g = pr.get("PageGUID", "")
            ps = pages.get(g, {})
            sx = ps.get("scale_x", 0) or 0
            if ps:
                page_names.add(ps["name"])
                scale_labels.add(ps["scale_label"])
            pts = _parse_points(pr.get("DigitizerData", ""))
            if cls == "Linear":
                if sx: total += _polyline_len(pts) / sx
            elif cls == "Area":
                if sx: total += _polygon_area(pts) / (sx * sx)
            elif cls == "Count":
                total += max(1, len(pts)) if pts else 1

        unit = {"Linear": "LF", "Area": "SF", "Count": "EA"}[cls]
        records.append({
            "name": name,
            "type": cls,
            "quantity": round(total, 2),
            "unit": unit,
            "pages": sorted(page_names),
            "scales": sorted(s for s in scale_labels if s),
            "n_sections": len(secs),
        })
    return {"pages": pages, "items": records}


# ── Item-name interpreter (learns scope + spec from the estimator's naming) ──
# These patterns are over the UNIVERSAL framing language, not architect-specific.
def interpret_name(name):
    """Pull structured meaning out of a takeoff item name."""
    out = {"raw": name}
    n = name.upper()

    # Building/area code prefix (N/S, CH, TE, MK, MC, GL, L2, L3, etc.)
    mb = re.match(r"\s*([A-Z]{1,3}(?:/[A-Z])?(?:&[A-Z])?)\s*[-–]", name)
    if mb: out["area_code"] = mb.group(1)

    # Level
    ml = re.search(r"\b(GROUND|GL|L1|L2|L3|LEVEL\s*\d)\b", n)
    if ml: out["level"] = ml.group(1)

    # Scope keywords
    if "GSF" in n: out["scope"] = "GSF area"
    elif "WALL" in n: out["scope"] = "wall framing"
    elif "LEDGER" in n: out["scope"] = "ledger"
    elif "FACIA" in n or "FASCIA" in n: out["scope"] = "fascia"
    elif "BLOCKING" in n: out["scope"] = "blocking"
    elif "WINDOW" in n: out["scope"] = "window"
    elif "UNIT" in n: out["scope"] = "unit count"
    elif re.search(r"\bH\d+-\d+", n): out["scope"] = "header"

    # Lumber spec
    ms = re.search(r"\(?\d?\)?\s*\d{1,2}X\d{1,2}", n)
    if ms: out["lumber"] = ms.group(0).strip()
    # OC spacing
    mo = re.search(r"(\d{1,2})\"?\s*O\.?C", n)
    if mo: out["oc"] = int(mo.group(1))
    # Header mark
    mh = re.search(r"\bH\d+-\d+F?\b", n)
    if mh: out["header_mark"] = mh.group(0)
    # Add/Remove (revision deltas)
    if re.search(r"\bADD\b", n): out["delta"] = "add"
    elif re.search(r"\bREMOVE\b", n): out["delta"] = "remove"

    return out


if __name__ == "__main__":
    import sys, json
    sj = sys.argv[1] if len(sys.argv) > 1 else \
         "/mnt/user-data/uploads/P0348_-_Willow_Way.SwiftJob"
    data = reconstruct_takeoff(sj)
    pages = data["pages"]
    items = data["items"]

    scaled_pages = [p for p in pages.values() if p["scale_x"]]
    print(f"Pages total: {len(pages)}  |  pages with scale set (measured): {len(scaled_pages)}")
    print(f"Takeoff items: {len(items)}\n")

    from collections import Counter
    by_type = Counter(i["type"] for i in items)
    print(f"By type: {dict(by_type)}\n")

    print("=== Reconstructed takeoff ===")
    for it in items:
        interp = interpret_name(it["name"])
        spec = " ".join(f"{k}={v}" for k, v in interp.items() if k != "raw")
        pg = it["pages"][0][:28] if it["pages"] else "-"
        sc = it["scales"][0] if it["scales"] else "-"
        print(f"  [{it['type'][:4]}] {it['name'][:46]:46s} = {it['quantity']:>9} {it['unit']}  @ {pg:28s} {sc}")