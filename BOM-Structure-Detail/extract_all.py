"""
extract_all.py  —  BOM Structure Detail extractor

Differences from the original scripts:
  - Processes ALL BOM files in each job folder (XML + By Structure PDF + Summary PDF)
    rather than picking only the "best" type.
  - Does NOT deduplicate by structure name across files.  If a remake, add-on, or
    revision BOM contains the same structure again, it is included as a separate row.
  - Adds columns: Source File, Source Subfolder, BOM Type
    so every row can be traced back to the exact file it came from.
  - Parameterized by year; run all four years from one script.

Usage:
    python extract_all.py 2026
    python extract_all.py 2025
    python extract_all.py 2024
    python extract_all.py 2023
    python extract_all.py          # runs all four years sequentially
"""

import csv
import os
import re
import sys
from datetime import date

from graph_client import acquire_token, ensure_fresh_token, GRAPH_ROOT, graph_get
from sharepoint_client import get_site, get_drive, list_children, download_file
from parse_bom_pdf import (
    parse_bom_pdf_safe,
    parse_bom_by_structure_pdf_safe,
    parse_bom_by_structure_xml,
    parse_shop_drawing_pdf_safe,
    parse_ncf_docx,
    normalize_location,
)

HOSTNAME = "armorockllc.sharepoint.com"

YEAR_CONFIG = {
    "2026": {"site_path": "/sites/JobData2026",  "drive_name": "Job Data 2026"},
    "2025": {"site_path": "/sites/jobdata2025",  "drive_name": "Job Data 2025"},
    "2024": {"site_path": "/sites/jobdata2024",  "drive_name": "Job Data 2024"},
    "2023": {"site_path": "/sites/jobdata2023",  "drive_name": "Job Data 2023"},
}

SKIP_FOLDERS = {"forms", "plugin_data", "robotinterface", "__macosx", "antiquated files"}

OUTPUT_DIR = "output"
SKIPPED_CSV       = os.path.join(OUTPUT_DIR, "skipped_structures.csv")
ERRORS_CSV        = os.path.join(OUTPUT_DIR, "errors.csv")
UNCLASSIFIED_CSV  = os.path.join(OUTPUT_DIR, "unclassified_files.csv")
_today = date.today()
DATE_EXTRACTED = f"{_today.month}/{_today.day}/{_today.year}"


def _write_skipped(job_code, filename, structure_name, year):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(SKIPPED_CSV, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["year", "job_code", "structure_name", "source_file"])
        w.writerow({"year": year, "job_code": job_code,
                    "structure_name": structure_name, "source_file": filename})


def _write_error(year, job_code, source_file, error_type, message):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(ERRORS_CSV, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["year", "job_code", "source_file", "error_type", "message"])
        w.writerow({"year": year, "job_code": job_code, "source_file": source_file,
                    "error_type": error_type, "message": str(message)})


_SKIP_EXTENSIONS = {".docx", ".doc", ".msg", ".eml", ".png", ".jpg", ".jpeg",
                    ".gif", ".bmp", ".tif", ".tiff", ".zip", ".7z", ".xlsx",
                    ".xls", ".txt", ".url", ".lnk", ".ini", ".db"}

def _write_unclassified(year, job_code, filename, subfolder):
    ext = os.path.splitext(filename)[1].lower()
    if ext in _SKIP_EXTENSIONS:
        return
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(UNCLASSIFIED_CSV, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["year", "job_code", "filename", "extension", "subfolder"])
        w.writerow({"year": year, "job_code": job_code, "filename": filename,
                    "extension": ext, "subfolder": subfolder})

_JOB_FOLDER_RE = re.compile(r"^([A-Z]{2,4})\s*[-–—û]\s*(.+)$")

FIELDNAMES = [
    "Year Release", "BOM Release Date", "Date extracted",
    "Job Code", "Project Name", "Structure Name",
    "Job Location", "Location Source", "Zip Code", "Contractor",
    "Agency", "Engineer",
    "Part Name", "Product Number", "Quantity", "Weight",
    "Production Part", "Part Type", "Part Subtype",
    "Source File", "Source Subfolder", "Source File Name",
]

# ── Part classification ────────────────────────────────────────────────────────

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
_PR_HDPE_SUBTYPE_MAP = [
    ("LADTECH", "Ladtech"), ("PRO-RING", "Pro-Ring"),
    ("GRADE", "Grade"), ("FINISH", "Finish"), ("FLAT", "Flat"),
]
_MHB_TROUGH_MAP = {
    ".133": '1.33"', ".15": "Flat Floor", ".75": "3/4 Depth",
    ".5": "1/2 Depth", ".1": '1"', "FF": "Flat Floor",
}
_MH_DIAMETERS = [192, 144, 120, 96, 84, 72, 60, 48]
_BOX_DIMS     = [154, 144, 120, 115, 104, 96, 91, 84, 79, 72, 65, 48, 36]


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


def classify_part(category, part_number, description):
    cat_lower = (category or "").lower()
    desc_up   = (description or "").upper()
    pn_up     = (part_number or "").upper()
    if "precast" in cat_lower and (pn_up.startswith("PR") or pn_up.startswith("HDPE")):
        subtype = next((v for k, v in _PR_HDPE_SUBTYPE_MAP if k in desc_up), "")
        return "Resale", subtype
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


# ── Location overrides ────────────────────────────────────────────────────────

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


# ── Document type detection (table-driven) ────────────────────────────────────

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


def _is_ncf(filename):
    n = filename.lower()
    return n.endswith(".docx") and (n.endswith("_ncf.docx") or n.endswith("_ecf.docx"))


def _filename_sort_date(filename):
    m = re.search(r"\b(\d{1,2})\.(\d{1,2})\.(\d{2})\b", filename)
    if m:
        return (2000 + int(m.group(3)), int(m.group(1)), int(m.group(2)))
    return (0, 0, 0)


# ── SharePoint traversal (yields file + relative subfolder path) ──────────────

def _iter_files_with_path(token, drive_id, folder_id, rel_path=""):
    from sharepoint_client import list_children
    for item in list_children(token, drive_id, folder_id):
        if "file" in item:
            yield item, rel_path
        elif "folder" in item:
            name = item["name"].lower()
            if name not in SKIP_FOLDERS:
                child_path = f"{rel_path}/{item['name']}" if rel_path else item["name"]
                yield from _iter_files_with_path(token, drive_id, item["id"], child_path)


# ── Row builder ───────────────────────────────────────────────────────────────

def _build_row(job_code, drive_year, header, structure_name, item,
               source_file, source_subfolder, bom_type, ncf_meta=None):
    part_type, subtype = classify_part(item["category"], item["part_number"], item["description"])
    pn_name = build_part_name(item["part_number"])
    if pn_name == (item.get("part_number") or "").strip():
        pn_name = item.get("description") or pn_name

    ncf = ncf_meta or {}
    bom_location = header.get("location", "")
    ncf_location = ncf.get("location", "")
    if bom_location:
        job_location    = bom_location
        location_source = "BOM"
    elif ncf_location:
        job_location    = ncf_location
        location_source = ncf.get("location_source", "NCF")
    else:
        job_location    = ""
        location_source = ""
    job_location = normalize_location(job_location)
    if not job_location and job_code in _LOCATION_OVERRIDES:
        ov = _LOCATION_OVERRIDES[job_code]
        job_location    = ov["location"]
        location_source = ov["source"]

    year_val = header.get("release_date", "")
    if year_val:
        parts = year_val.split("/")
        year_val_out = parts[2] if len(parts) == 3 and parts[2].isdigit() else drive_year
    else:
        year_val_out = drive_year

    return {
        "Year Release":     year_val_out,
        "BOM Release Date": header.get("release_date", ""),
        "Date extracted":   DATE_EXTRACTED,
        "Job Code":         job_code,
        "Project Name":     header.get("job_name", "") or ncf.get("job_name", ""),
        "Structure Name":   structure_name,
        "Job Location":     job_location,
        "Location Source":  location_source,
        "Zip Code":         ncf.get("zipcode", ""),
        "Contractor":       header.get("contractor", "") or ncf.get("customer", ""),
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
        "Source Subfolder": source_subfolder,
        "Source File Name": bom_type,
    }


def _parse_summary_structure(filename, pdf_job_name):
    base = re.sub(r"-BOM Summary\.pdf$", "", filename, flags=re.IGNORECASE)
    rest = re.sub(r"^[A-Z]{3}\s+\d+\.\d+\.\d{2}\s+", "", base)
    clean_job = re.sub(r"^[A-Z]{3}\s+", "", pdf_job_name or "").strip()
    if clean_job and rest.startswith(clean_job):
        rest = rest[len(clean_job):].lstrip("-").strip()
    else:
        rest = rest.strip()
    return rest or pdf_job_name or ""


# ── Per-job processor ─────────────────────────────────────────────────────────

def process_job(token, drive_id, drive_year, job_code, folder_id, use_ncf=True):
    all_files = list(_iter_files_with_path(token, drive_id, folder_id))

    bom_files = [(f, path) for f, path in all_files if classify_document(f["name"])]
    ncf_files = [f for f, _ in all_files if _is_ncf(f["name"])]
    if not ncf_files and use_ncf:
        ncf_files = [f for f, _ in all_files if f["name"].lower().endswith(".docx")]

    classified_ids = {f["id"] for f, _ in bom_files} | {f["id"] for f in ncf_files}
    for f, path in all_files:
        if f["id"] not in classified_ids:
            _write_unclassified(drive_year, job_code, f["name"], path)

    if not bom_files:
        print(f"  [{job_code}] No BOM files found — skipping")
        return []

    # If folder name wasn't a standard job code, try to pull it from the BOM filename
    if not re.match(r'^[A-Z]{2,4}$', job_code):
        for f, _ in bom_files:
            m = re.match(r'^([A-Z]{2,4})\s', f["name"])
            if m:
                job_code = m.group(1)
                break

    # If still unresolved, check the first component of the subfolder path
    # (handles category folders like _PC that contain real job subfolders e.g. CVA- Canterwood Manholes)
    if not re.match(r'^[A-Z]{2,4}$', job_code):
        for _, subfolder in bom_files:
            if subfolder:
                first_component = subfolder.split('/')[0]
                m = _JOB_FOLDER_RE.match(first_component.strip())
                if m:
                    job_code = m.group(1)
                    break

    job_code_resolved = bool(re.match(r'^[A-Z]{2,4}$', job_code))

    ncf_meta = None
    if use_ncf:
        for ncf_file in ncf_files:
            try:
                ncf_bytes = download_file(token, drive_id, ncf_file["id"])
                candidate = parse_ncf_docx(ncf_bytes)
                if ncf_meta is None:
                    ncf_meta = candidate
                if candidate.get("location"):
                    ncf_meta = candidate
                    break
            except Exception as exc:
                print(f"  [{job_code}] NCF error ({ncf_file['name']}): {exc}")
                _write_error(drive_year, job_code, ncf_file["name"], "NCF error", exc)

    bom_files.sort(key=lambda t: (
        _LABEL_ORDER.get(classify_document(t[0]["name"]), 99),
        tuple(-x for x in _filename_sort_date(t[0]["name"])),
    ))

    rows = []
    for file_item, subfolder in bom_files:
        fname = file_item["name"]
        label = classify_document(fname)
        print(f"  [{job_code}] ({label}) {subfolder or '.'}/{fname}")
        try:
            file_bytes = download_file(token, drive_id, file_item["id"])
            if label == "BOM by Structure XML":
                bom = parse_bom_by_structure_xml(file_bytes)
            elif label == "BOM by Structure PDF":
                bom = parse_bom_by_structure_pdf_safe(file_bytes)
            elif label == "BOM Summary PDF":
                bom = parse_bom_pdf_safe(file_bytes)
            elif label == "Shop Drawing PDF":
                bom = parse_shop_drawing_pdf_safe(file_bytes)
            else:
                continue

            for skipped in bom.get("skipped_structures", []):
                print(f'  [{job_code}] [SKIP structure] "{skipped}" in {fname} — no pattern matched')
                _write_skipped(job_code, fname, skipped, year)

            # If job_code still unresolved, try BOM header when structures are present
            if not job_code_resolved and bom.get("structures"):
                for candidate in [bom["header"].get("job_number", ""), bom["header"].get("job_name", "")]:
                    m = re.match(r'^([A-Z]{2,4})(?=[^A-Z]|$)', candidate.strip().upper())
                    if m:
                        new_code = m.group(1)
                        for r in rows:
                            r["Job Code"] = new_code
                        job_code = new_code
                        job_code_resolved = True
                        print(f"  [{job_code}] job code resolved from BOM content")
                        break

            if label in ("BOM by Structure XML", "BOM by Structure PDF", "Shop Drawing PDF"):
                for structure in bom["structures"]:
                    sname = structure["structure_name"]
                    for item in structure["line_items"]:
                        rows.append(_build_row(job_code, drive_year, bom["header"], sname,
                                               item, fname, subfolder, label, ncf_meta))
            else:  # BOM Summary PDF
                sname = _parse_summary_structure(fname, bom["header"].get("job_name", ""))
                for item in bom["line_items"]:
                    rows.append(_build_row(job_code, drive_year, bom["header"], sname,
                                           item, fname, subfolder, label, ncf_meta))
        except BaseException as exc:
            print(f"  [{job_code}] ERROR parsing {fname}: {type(exc).__name__}: {exc}")
            _write_error(drive_year, job_code, fname, type(exc).__name__, exc)
            if isinstance(exc, KeyboardInterrupt):
                raise
    return rows


# ── Year runner ───────────────────────────────────────────────────────────────

def run_year(token, year):
    cfg = YEAR_CONFIG[year]
    use_ncf = year != "2023"
    output_path = os.path.join(OUTPUT_DIR, f"all_bom_{year}.csv")

    print(f"\n{'='*60}")
    print(f"Year: {year}  site={cfg['site_path']}  drive={cfg['drive_name']}")

    site  = get_site(token, HOSTNAME, cfg["site_path"])
    drive = get_drive(token, site["id"], cfg["drive_name"])
    print(f"Drive: {drive['name']}")

    children = list_children(token, drive["id"], "root")
    job_folders = [c for c in children if "folder" in c]
    print(f"{len(job_folders)} job folders")

    all_rows = []
    errors   = []

    for folder in job_folders:
        token = ensure_fresh_token() or token
        m = _JOB_FOLDER_RE.match(folder["name"].strip())
        job_code = m.group(1) if m else folder["name"].strip()
        try:
            rows = process_job(token, drive["id"], year, job_code, folder["id"], use_ncf)
            all_rows.extend(rows)
        except KeyboardInterrupt:
            raise
        except Exception as exc:
            msg = f"[{job_code}] FOLDER ERROR: {type(exc).__name__}: {exc}"
            print(f"  {msg}")
            errors.append(msg)
            _write_error(year, job_code, folder["name"], "FOLDER ERROR", exc)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    if all_rows:
        with open(output_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
            writer.writeheader()
            writer.writerows(all_rows)
        print(f"\nWrote {len(all_rows):,} rows -> {output_path}")
    else:
        print(f"\n(no rows for {year})")

    if errors:
        print(f"{len(errors)} errors during {year} run")

    return all_rows


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    years = sys.argv[1:] if len(sys.argv) > 1 else list(YEAR_CONFIG.keys())
    for y in years:
        if y not in YEAR_CONFIG:
            print(f"Unknown year: {y}  (valid: {list(YEAR_CONFIG.keys())})")
            sys.exit(1)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    for path, fields in [
        (SKIPPED_CSV,      ["year", "job_code", "structure_name", "source_file"]),
        (ERRORS_CSV,       ["year", "job_code", "source_file", "error_type", "message"]),
        (UNCLASSIFIED_CSV, ["year", "job_code", "filename", "extension", "subfolder"]),
    ]:
        with open(path, "w", newline="", encoding="utf-8") as f:
            csv.DictWriter(f, fieldnames=fields).writeheader()

    token = acquire_token()
    for y in years:
        run_year(token, y)


if __name__ == "__main__":
    main()
