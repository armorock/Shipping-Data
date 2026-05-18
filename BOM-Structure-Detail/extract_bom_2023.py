import csv
import os
import re
import sys
from datetime import date
from graph_client import acquire_token, GRAPH_ROOT, graph_get, graph_get_all
from sharepoint_client import get_site, get_drive, list_children, iter_files, download_file
from parse_bom_pdf import (
    parse_bom_pdf_safe, parse_bom_by_structure_pdf_safe,
    parse_bom_by_structure_xml, parse_shop_drawing_pdf_safe,
    normalize_location,
)

HOSTNAME   = "armorockllc.sharepoint.com"
SITE_PATH  = "/sites/jobdata2023"
DRIVE_NAME = "Job Data 2023"
JOBS_ROOT  = "root"
OUTPUT_DIR = "output"
OUTPUT     = os.path.join(OUTPUT_DIR, "bom_manhole_map_2023.csv")

_drive_year_m = re.search(r"\b(\d{4})\b", DRIVE_NAME)
DRIVE_YEAR = _drive_year_m.group(1) if _drive_year_m else ""

_today = date.today()
DATE_EXTRACTED = f"{_today.month}/{_today.day}/{_today.year}"

_JOB_FOLDER_RE = re.compile(r"^([A-Z]{2,4})\s*[-–—û]\s*(.+)$")

def _load_location_overrides():
    path = os.path.join("data", "location_overrides.csv")
    if not os.path.exists(path):
        return {}
    overrides = {}
    with open(path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            overrides[row["Job Code"]] = {"location": row["Job Location"], "source": row["Location Source"]}
    return overrides

_LOCATION_OVERRIDES = _load_location_overrides()

_PRECAST_TYPE_MAP = [
    ("MHGR", "Grade Ring"), ("MHTL", "Taper Lid"), ("MHHF", "Flat Top"),
    ("MHTE", "Tee"), ("MHRE", "Reducer"), ("MHL", "Lid"),
    ("MHS", "Section"), ("MHB", "Base"), ("MHC", "Cone"),
    ("RMH", "Rehab Ring"), ("BOX", "Box Culvert"),
]
_PRECAST_SUBTYPE_MAP = [
    ("NO BELL", "No Bell"), ("ECCENTRIC", "Eccentric"), ("ECC POLYMER", "Eccentric"),
    ("ECC", "Eccentric"), ("FLAT", "Flat"), ("CONCENTRIC", "Concentric"),
]
_RESALE_SUBTYPE_MAP = [
    ("MASTIC WRAP", "Mastic Wrap"), ("MASTIC", "Mastic"), ("GASKET", "Gasket"),
    ("GROUT", "Grout"), ("EPOXY", "Epoxy"), ("SHACKLE", "Hardware"),
]

_MHB_TROUGH_MAP = {
    ".133": '1.33"', ".15": "Flat Floor", ".75": "3/4 Depth",
    ".5": "1/2 Depth", ".1": '1"', "FF": "Flat Floor",
}
_MH_DIAMETERS = [192, 144, 120, 96, 84, 72, 60, 48]
_BOX_DIMS     = [154, 144, 120, 115, 104, 96, 91, 84, 79, 72, 65, 48, 36]


def _find_2023_site(token):
    """Search Graph for SharePoint sites containing '2023' and print candidates."""
    print("Searching for SharePoint sites containing '2023'...")
    try:
        results = graph_get(token, f"{GRAPH_ROOT}/sites", {"search": "2023"})
        sites = results.get("value", [])
        if not sites:
            print("  No sites found matching '2023'.")
        else:
            print(f"  Found {len(sites)} candidate site(s):")
            for s in sites:
                name  = s.get("displayName", "")
                wurl  = s.get("webUrl", "")
                sid   = s.get("id", "")
                parts = sid.split(",")
                path  = "/" + wurl.split(HOSTNAME, 1)[-1].lstrip("/") if HOSTNAME in wurl else wurl
                print(f"    displayName: {name}")
                print(f"    webUrl:      {wurl}")
                print(f"    SITE_PATH:   {path}")
                print()
    except Exception as exc:
        print(f"  Site search failed: {exc}")


def _parse_dia(s, known):
    for d in known:
        ds = str(d)
        if s.startswith(ds):
            return ds, s[len(ds):]
    return None, s


def _parse_box_dims(s):
    dims = []
    while s and s[0].isdigit():
        dia, rest = _parse_dia(s, _BOX_DIMS)
        if dia:
            dims.append(dia)
            s = rest
        else:
            m = re.match(r"(\d+)", s)
            dims.append(m.group(1))
            s = s[m.end():]
            break
    return dims


def build_part_name(product_number):
    pn = (product_number or "").strip().upper()
    if not pn:
        return ""

    m = re.match(r"^MHGRB(\d{2,3})", pn)
    if m:
        return f'{m.group(1)}" Grade Ring Barrel'
    m = re.match(r"^MHTLC(\d{2,3})(\d{2})", pn)
    if m:
        return f'{m.group(1)}" Top Lid Collar {int(m.group(2))}" Opening'
    m = re.match(r"^RMHLC(\d{2})(\d{2})", pn)
    if m:
        return f'{m.group(1)}" Rehab Lid Collar {int(m.group(2))}" Opening'
    m = re.match(r"^MHCC(\d{2})(\d{2})(-[235])?", pn)
    if m:
        return f'{m.group(1)}" Conc Cone {int(m.group(2))}"'
    m = re.match(r"^MHGR(\d{2})X(\d)", pn)
    if m:
        return f'{m.group(1)}" Grade Ring x{m.group(2)}'
    m = re.match(r"^MHLC", pn)
    if m:
        rest = pn[4:]
        dia, rest = _parse_dia(rest, _MH_DIAMETERS)
        if dia:
            if rest.startswith("HATCH"):
                return f'{dia}" Lid Collar Hatch'
            m2 = re.match(r"(\d{2})", rest)
            opening = m2.group(1) if m2 else None
            return f'{dia}" Lid Collar {int(opening)}" Opening' if opening else f'{dia}" Lid Collar'
    m = re.match(r"^MHTL(\d{2,3})(\d{2})(-[\d]+)?", pn)
    if m:
        return f'{m.group(1)}" Top Lid {int(m.group(2))}" Opening'
    m = re.match(r"^RMHC(\d{2})(\d{2})", pn)
    if m:
        return f'{m.group(1)}" Rehab Cone {int(m.group(2))}"'
    m = re.match(r"^RMHL(\d{2})(\d{2})?(CAST|HATCH)?", pn)
    if m:
        dia, opening, sfx = m.groups()
        if sfx == "HATCH":
            return f'{dia}" Rehab Lid Hatch'
        return f'{dia}" Rehab Lid {int(opening)}" Opening' if opening else f'{dia}" Rehab Lid'
    m = re.match(r"^BOX([FLST])(\d.*)", pn)
    if m:
        sub, dim_str = m.group(1), m.group(2)
        sfx = ""
        if dim_str.endswith("HATCH"):
            dim_str, sfx = dim_str[:-5], " Hatch"
        elif dim_str.endswith("FF"):
            dim_str, sfx = dim_str[:-2], " Flat Floor"
        sub_map = {"F": "Flat Floor Box", "L": "Box Lid", "S": "Box Section", "T": "Box"}
        dims = _parse_box_dims(dim_str)
        return sub_map[sub] + (" " + "x".join(dims) if dims else "") + sfx
    m = re.match(r"^MHB(\d{2,3})(\d{2,3})(\.133|\.15|\.75|\.5|\.1|FF)?(\.?ES)?(/DE2|DE2|/DE|DE)?(-[235])?", pn)
    if m:
        dia, ht, trough, es, de, _ = m.groups()
        parts = [f'{dia}"', "Base", f'{int(ht)}"']
        if trough:
            parts.append(_MHB_TROUGH_MAP.get(trough, trough.lstrip(".")))
        if es:
            parts.append("ES")
        if de:
            parts.append("DE2" if "2" in de else "DE")
        return " ".join(parts)
    m = re.match(r"^MHC(\d{2})(\d{2})(-[235])?", pn)
    if m:
        return f'{m.group(1)}" Cone {int(m.group(2))}"'
    m = re.match(r"^MHL", pn)
    if m:
        rest = pn[3:]
        if rest.startswith("HATCH"):
            return "Lid Hatch"
        dia, rest = _parse_dia(rest, _MH_DIAMETERS)
        if dia:
            if rest.startswith("HATCH"):
                return f'{dia}" Lid Hatch'
            m2 = re.match(r"(\d{2})", rest)
            opening = m2.group(1) if m2 else None
            return f'{dia}" Lid {int(opening)}" Opening' if opening else f'{dia}" Lid'
    m = re.match(r"^MHS(\d{2,3})(\d{2,3})(NBH|NB|H|L)?(-[235L])?", pn)
    if m:
        dia, ht, sfx, _ = m.groups()
        sfx_map = {"NB": "NB", "H": "Hole", "NBH": "NB Hole", "L": "L"}
        name = f'{dia}" Section {int(ht)}"'
        if sfx:
            name += f' {sfx_map.get(sfx, sfx)}'
        return name
    m = re.match(r"^MHT(\d{3})-([\d.]+)", pn)
    if m:
        return f'{m.group(1)}" Top {m.group(2)}'
    m = re.match(r"^RMH(\d{2})(\d{2})(H|-?L)?", pn)
    if m:
        dia, ht, sfx = m.groups()
        name = f'{dia}" Rehab Riser {int(ht)}"'
        if sfx == "H":
            name += " Hole"
        return name
    m = re.match(r"^BOX(\d.*)", pn)
    if m:
        dim_str = m.group(1)
        sfx = ""
        if dim_str.endswith("HATCH"):
            dim_str, sfx = dim_str[:-5], " Hatch"
        elif dim_str.endswith("FF"):
            dim_str, sfx = dim_str[:-2], " Flat Floor"
        dims = _parse_box_dims(dim_str)
        return "Box" + (" " + "x".join(dims) if dims else "") + sfx
    m = re.match(r"^MT(\d+)", pn)
    if m:
        return f'{m.group(1)}" Misc Top'

    return product_number


def parse_folder_name(name):
    m = _JOB_FOLDER_RE.match(name.strip())
    if m:
        return m.group(1), m.group(2).strip()
    return None, name.strip()


def _load_document_types():
    path = os.path.join("data", "document_types.csv")
    if not os.path.exists(path):
        return []
    rows = []
    with open(path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            rows.append((row["keyword"], row["extension"], row["source_file_name"]))
    return rows

_DOC_TYPES = _load_document_types()

_LABEL_ORDER = {
    "BOM by Structure XML": 0,
    "BOM by Structure PDF": 1,
    "BOM Summary PDF":      2,
    "Shop Drawing PDF":     3,
    "Quotation PDF":        4,
}


def classify_document(filename):
    n = filename.lower()
    ext = os.path.splitext(n)[1]
    for keyword, extension, label in _DOC_TYPES:
        if keyword in n and ext == extension:
            return label
    return None


def _filename_sort_date(filename):
    m = re.search(r"\b(\d{1,2})\.(\d{1,2})\.(\d{2})\b", filename)
    if m:
        return (2000 + int(m.group(3)), int(m.group(1)), int(m.group(2)))
    return (0, 0, 0)


def _year_from_release_date(release_date, filename_fallback=""):
    if release_date:
        parts = release_date.split("/")
        if len(parts) == 3 and parts[2].isdigit():
            return parts[2]
    m = re.search(r"\b\d{1,2}\.\d{1,2}\.(\d{2})\b", filename_fallback)
    return f"20{m.group(1)}" if m else ""


def parse_structure_from_summary_filename(filename, pdf_job_name):
    base = re.sub(r"-BOM Summary\.pdf$", "", filename, flags=re.IGNORECASE)
    rest = re.sub(r"^[A-Z]{2,4}\s+\d+\.\d+\.\d{2}\s+", "", base)
    clean_job = re.sub(r"^[A-Z]{2,4}\s+", "", pdf_job_name).strip()
    if clean_job and rest.startswith(clean_job):
        return rest[len(clean_job):].lstrip("-").strip()
    return ""


def classify_part(category, part_number, description):
    cat_lower = (category or "").lower()
    desc_up   = (description or "").upper()
    pn_up     = (part_number or "").upper()

    if "precast" in cat_lower:
        part_type = next((v for k, v in _PRECAST_TYPE_MAP if pn_up.startswith(k)), "")
        if not part_type:
            for kw, v in [("CONE","Cone"),("LID","Lid"),("SECTION","Section"),
                          ("BASE","Base"),("GRADE RING","Grade Ring"),("RISER","Riser")]:
                if kw in desc_up:
                    part_type = v
                    break
        subtype = next((v for k, v in _PRECAST_SUBTYPE_MAP if k in desc_up), "")
        return part_type, subtype
    else:
        subtype = next((v for k, v in _RESALE_SUBTYPE_MAP if k in desc_up), "")
        return "Resale", subtype


def build_row(job_code, header, structure_name, year, item,
              source_file="", source_file_name=""):
    part_type, subtype = classify_part(
        item["category"], item["part_number"], item["description"]
    )
    pn_name = build_part_name(item["part_number"])
    if pn_name == (item.get("part_number") or "").strip():
        pn_name = item.get("description") or pn_name
    job_location    = normalize_location(header["location"])
    location_source = "BOM" if job_location else ""
    if not job_location and job_code in _LOCATION_OVERRIDES:
        ov = _LOCATION_OVERRIDES[job_code]
        job_location    = ov["location"]
        location_source = ov["source"]
    return {
        "Year Release":     year,
        "BOM Release Date": header.get("release_date", ""),
        "Date extracted":   DATE_EXTRACTED,
        "Job Code":         job_code,
        "Project Name":     header["job_name"],
        "Structure Name":   structure_name,
        "Job Location":     job_location,
        "Location Source":  location_source,
        "Contractor":       header["contractor"],
        "Agency":           "",
        "Engineer":         "",
        "Part Name":        pn_name,
        "Product Number":   item["part_number"],
        "Quantity":         item["quantity"],
        "Weight":           item["weight_lbs"],
        "Production Part":  item["category"],
        "Part Type":        part_type,
        "Part Subtype":     subtype,
        "Source File":      source_file,
        "Source Subfolder": "",
        "Source File Name": source_file_name,
    }


def _process_by_structure_bom(job_code, file_item, bom, label, seen_structures):
    rows = []
    for skipped in bom.get("skipped_structures", []):
        print(f'[{job_code}] [SKIP structure] "{skipped}" in {file_item["name"]} — no pattern matched')
    for structure in bom["structures"]:
        name = structure["structure_name"]
        if name in seen_structures:
            continue
        seen_structures.add(name)
        for item in structure["line_items"]:
            rows.append(build_row(job_code, bom["header"], name, DRIVE_YEAR, item,
                                  file_item["name"], label))
    return rows


def process_by_structure_xml(job_code, file_item, xml_bytes, seen_structures):
    bom = parse_bom_by_structure_xml(xml_bytes)
    return _process_by_structure_bom(job_code, file_item, bom,
                                     "BOM by Structure XML", seen_structures)


def process_by_structure(job_code, file_item, pdf_bytes, seen_structures):
    bom = parse_bom_by_structure_pdf_safe(pdf_bytes)
    return _process_by_structure_bom(job_code, file_item, bom,
                                     "BOM by Structure PDF", seen_structures)


def process_shop_drawing(job_code, file_item, pdf_bytes, seen_structures):
    bom = parse_shop_drawing_pdf_safe(pdf_bytes)
    return _process_by_structure_bom(job_code, file_item, bom,
                                     "Shop Drawing PDF", seen_structures)


def process_summary(job_code, file_item, pdf_bytes, seen_structures):
    bom = parse_bom_pdf_safe(pdf_bytes)
    structure = parse_structure_from_summary_filename(file_item["name"], bom["header"]["job_name"])
    if structure in seen_structures:
        return []
    seen_structures.add(structure)
    rows = []
    for item in bom["line_items"]:
        rows.append(build_row(job_code, bom["header"], structure, DRIVE_YEAR, item,
                              file_item["name"], "BOM Summary PDF"))
    return rows


def write_csv(path, rows):
    if not rows:
        print(f"  (no rows for {path})")
        return
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)
    print(f"Wrote {len(rows):,} rows -> {path}")


def main():
    token = acquire_token()

    try:
        site = get_site(token, HOSTNAME, SITE_PATH)
    except Exception as exc:
        print(f"Could not find site at {SITE_PATH}: {exc}")
        print()
        _find_2023_site(token)
        print()
        print("Update SITE_PATH and DRIVE_NAME at the top of this script, then re-run.")
        sys.exit(1)

    print(f"Site: {site['displayName']}")

    try:
        drive = get_drive(token, site["id"], DRIVE_NAME)
    except ValueError as exc:
        print(f"Could not find drive '{DRIVE_NAME}': {exc}")
        sys.exit(1)

    print(f"Drive: {drive['name']}")

    children = list_children(token, drive["id"], JOBS_ROOT)
    job_folders = [c for c in children if "folder" in c]
    print(f"Found {len(job_folders)} top-level folders")
    for f in job_folders:
        code, _ = parse_folder_name(f["name"])
        tag = code if code else "[no-match]"
        print(f"  {tag:>10}  {f['name']}")
    print()

    all_rows = []
    skipped  = []
    errors   = []

    for folder in job_folders:
        job_code, _ = parse_folder_name(folder["name"])
        if not job_code:
            skipped.append(folder["name"])
            continue

        try:
            all_files    = list(iter_files(token, drive["id"], folder["id"], recursive=True))
            classified = [(f, classify_document(f["name"])) for f in all_files]
            bom_files_all = [(f, lbl) for f, lbl in classified if lbl]
            bom_files_all.sort(key=lambda t: (
                _LABEL_ORDER.get(t[1], 99),
                tuple(-x for x in _filename_sort_date(t[0]["name"])),
            ))

            best_label = bom_files_all[0][1] if bom_files_all else None
            if not best_label:
                print(f"[{job_code}] No BOM file — skipping")
                continue
            bom_files = [f for f, lbl in bom_files_all if lbl == best_label]

            seen_structures = set()
            for file_item in bom_files:
                label = classify_document(file_item["name"])
                print(f"[{job_code}] ({label}) {file_item['name']}")
                try:
                    file_bytes = download_file(token, drive["id"], file_item["id"])
                    if label == "BOM by Structure XML":
                        rows = process_by_structure_xml(job_code, file_item, file_bytes, seen_structures)
                    elif label == "BOM by Structure PDF":
                        rows = process_by_structure(job_code, file_item, file_bytes, seen_structures)
                    elif label == "Shop Drawing PDF":
                        rows = process_shop_drawing(job_code, file_item, file_bytes, seen_structures)
                    else:
                        rows = process_summary(job_code, file_item, file_bytes, seen_structures)
                    all_rows.extend(rows)
                except BaseException as exc:
                    msg = f"[{job_code}] ERROR {file_item['name']}: {type(exc).__name__}: {exc}"
                    print(f"  {msg}", flush=True)
                    errors.append(msg)
                    if isinstance(exc, KeyboardInterrupt):
                        raise

        except BaseException as exc:
            msg = f"[{job_code}] FOLDER ERROR: {type(exc).__name__}: {exc}"
            print(msg, flush=True)
            errors.append(msg)
            if isinstance(exc, KeyboardInterrupt):
                raise

    print()
    write_csv(OUTPUT, all_rows)

    if skipped:
        n = len(skipped)
        print(f"\nSkipped {n} folders: {skipped[:5]}{'...' if n > 5 else ''}")
    if errors:
        print(f"\n{len(errors)} errors logged above")


if __name__ == "__main__":
    main()
