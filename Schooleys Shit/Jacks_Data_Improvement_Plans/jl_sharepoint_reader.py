"""
Traverses all 4 SharePoint year sites (2026–2023), parses every job folder's
documents (BOM XML/PDF, NCF/ECF docx, Shop Drawing PDF), and produces:

  output/jobcode_db_sharepoint.json
    One record per job code with cross-document comparison fields:
    - location from BOM vs location from NCF (conflict flag + detail)
    - contractor from BOM vs customer from NCF (match flag)
    - structure count from BOM vs structure count from Shop Drawing (conflict flag)
    - which documents were found / missing per job

First run: if the OAuth token is expired you'll be prompted to open a URL
in your browser and enter a device code (one-time, takes ~30 seconds).

Run: python jl_sharepoint_reader.py
"""

import os
import re
import sys
import json

_BOM_TOOLS = os.path.abspath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "BOM Structure Detail")
)
sys.path.insert(0, _BOM_TOOLS)

from graph_client import acquire_token, ensure_fresh_token
from sharepoint_client import get_site, get_drive, list_children, download_file
from parse_bom_pdf import (
    parse_bom_by_structure_xml,
    parse_bom_by_structure_pdf_safe,
    parse_bom_pdf_safe,
    parse_shop_drawing_pdf_safe,
    parse_ncf_docx,
    normalize_location,
)

BASE       = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE, "output")
OUT_PATH   = os.path.join(OUTPUT_DIR, "jobcode_db_sharepoint.json")

HOSTNAME = "armorockllc.sharepoint.com"
YEAR_CONFIG = {
    "2026": {"site_path": "/sites/JobData2026",  "drive_name": "Job Data 2026"},
    "2025": {"site_path": "/sites/jobdata2025",  "drive_name": "Job Data 2025"},
    "2024": {"site_path": "/sites/jobdata2024",  "drive_name": "Job Data 2024"},
    "2023": {"site_path": "/sites/jobdata2023",  "drive_name": "Job Data 2023"},
}

SKIP_FOLDERS = {"forms", "plugin_data", "robotinterface", "__macosx", "antiquated files"}

_JOB_FOLDER_RE = re.compile(r"^([A-Z]{2,4})\s*[-–—û]\s*(.+)$")

# Mirror of BOM Structure Detail/data/document_types.csv
_DOC_KEYWORDS = [
    ("bom by structure", ".xml",  "BOM by Structure XML"),
    ("bom by structure", ".pdf",  "BOM by Structure PDF"),
    ("bom summary",      ".pdf",  "BOM Summary PDF"),
    ("shop drawing",     ".pdf",  "Shop Drawing PDF"),
    ("shop drawings",    ".pdf",  "Shop Drawing PDF"),
]
_DOC_LABEL_ORDER = {
    "BOM by Structure XML": 0,
    "BOM by Structure PDF": 1,
    "BOM Summary PDF":      2,
    "Shop Drawing PDF":     3,
}


def _classify_doc(filename):
    n = filename.lower()
    ext = os.path.splitext(n)[1]
    for keyword, extension, label in _DOC_KEYWORDS:
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


def _iter_folder(token, drive_id, folder_id, rel_path=""):
    for item in list_children(token, drive_id, folder_id):
        if "file" in item:
            yield item, rel_path
        elif "folder" in item:
            name = item["name"].lower()
            if name not in SKIP_FOLDERS:
                child = f"{rel_path}/{item['name']}" if rel_path else item["name"]
                yield from _iter_folder(token, drive_id, item["id"], child)


def _safe_xml(data):
    try:
        return parse_bom_by_structure_xml(data)
    except Exception:
        return None


def _safe_ncf(data):
    try:
        return parse_ncf_docx(data)
    except Exception:
        return None


def _parse_city_state(loc_str):
    if not loc_str:
        return "", ""
    parts = loc_str.rsplit(",", 1)
    if len(parts) == 2:
        return parts[0].strip(), parts[1].strip()
    return loc_str.strip(), ""


def process_job(token, drive_id, year, job_code, folder_id, use_ncf=True):
    """Download and parse all documents in a job folder. Returns a raw dict."""
    all_files = list(_iter_folder(token, drive_id, folder_id))

    bom_files  = [(f, p) for f, p in all_files if _classify_doc(f["name"])]
    ncf_candidates = [f for f, _ in all_files if _is_ncf(f["name"])]
    if not ncf_candidates and use_ncf:
        ncf_candidates = [f for f, _ in all_files if f["name"].lower().endswith(".docx")]

    bom_files.sort(key=lambda t: (
        _DOC_LABEL_ORDER.get(_classify_doc(t[0]["name"]), 99),
        tuple(-x for x in _filename_sort_date(t[0]["name"])),
    ))

    result = {
        "job_code":             job_code,
        "sharepoint_year":      year,
        "bom_header":           None,
        "bom_structure_count":  None,
        "bom_file":             None,
        "bom_doc_type":         None,
        "ncf_meta":             None,
        "ncf_file":             None,
        "shop_structure_count": None,
        "shop_drawing_file":    None,
        "documents_found":      [],
        "parse_errors":         [],
    }

    # Parse NCF first so BOM can override location if both present
    if use_ncf:
        for ncf_item in ncf_candidates:
            try:
                data = download_file(token, drive_id, ncf_item["id"])
                parsed = _safe_ncf(data)
                if parsed:
                    if result["ncf_meta"] is None or parsed.get("location"):
                        result["ncf_meta"] = parsed
                        result["ncf_file"] = ncf_item["name"]
                    if parsed.get("location"):
                        break
            except Exception as e:
                result["parse_errors"].append(f"NCF({ncf_item['name']}): {e}")
        if result["ncf_meta"] is not None:
            result["documents_found"].append("NCF")

    # Parse BOM and Shop Drawing docs
    for file_item, _ in bom_files:
        fname = file_item["name"]
        label = _classify_doc(fname)
        try:
            data = download_file(token, drive_id, file_item["id"])
            if label == "BOM by Structure XML":
                parsed = _safe_xml(data)
            elif label == "BOM by Structure PDF":
                parsed = parse_bom_by_structure_pdf_safe(data)
            elif label == "BOM Summary PDF":
                parsed = parse_bom_pdf_safe(data)
            elif label == "Shop Drawing PDF":
                parsed = parse_shop_drawing_pdf_safe(data)
            else:
                continue

            if not parsed:
                continue

            if label == "Shop Drawing PDF":
                if result["shop_structure_count"] is None:
                    result["shop_structure_count"] = len(parsed.get("structures", []))
                    result["shop_drawing_file"] = fname
                    result["documents_found"].append("Shop Drawing PDF")
            else:
                if result["bom_header"] is None:
                    result["bom_header"] = parsed.get("header", {})
                    result["bom_structure_count"] = len(parsed.get("structures", [])) if "structures" in parsed else None
                    result["bom_file"] = fname
                    result["bom_doc_type"] = label
                    result["documents_found"].append(label)

        except Exception as e:
            result["parse_errors"].append(f"{label}({fname}): {e}")

    return result


def build_record(raw):
    """Flatten raw parsed data into a single record with conflict fields."""
    bom = raw.get("bom_header") or {}
    ncf = raw.get("ncf_meta") or {}

    project_name = bom.get("job_name") or ncf.get("job_name") or ""
    project_name_source = "BOM" if bom.get("job_name") else ("NCF" if ncf.get("job_name") else "")

    release_date = bom.get("release_date") or ""
    year_released = ""
    if release_date:
        parts = release_date.split("/")
        if len(parts) == 3 and len(parts[2]) == 4 and parts[2].isdigit():
            year_released = parts[2]
        elif len(parts) == 3 and len(parts[2]) == 2 and parts[2].isdigit():
            year_released = str(2000 + int(parts[2]))
    if not year_released:
        year_released = raw.get("sharepoint_year") or ""

    loc_bom = normalize_location(bom.get("location") or "")
    loc_ncf = normalize_location(ncf.get("location") or "")

    loc_conflict = bool(loc_bom and loc_ncf and loc_bom != loc_ncf)
    loc_conflict_detail = f"BOM='{loc_bom}'; NCF='{loc_ncf}'" if loc_conflict else None

    location = loc_bom or loc_ncf or ""
    location_source = "BOM" if loc_bom else ("NCF" if loc_ncf else "")
    city, state = _parse_city_state(location)

    contractor_bom = bom.get("contractor") or ""
    customer_ncf   = ncf.get("customer") or ""

    def _norm(s):
        return re.sub(r"\s+", " ", (s or "").strip().lower())

    contractor_customer_match = None
    if contractor_bom and customer_ncf:
        contractor_customer_match = _norm(contractor_bom) == _norm(customer_ncf)

    sc_bom  = raw.get("bom_structure_count")
    sc_shop = raw.get("shop_structure_count")
    sc_conflict = bool(sc_bom is not None and sc_shop is not None and sc_bom != sc_shop)
    sc_conflict_detail = f"BOM={sc_bom}; ShopDrawing={sc_shop}" if sc_conflict else None

    found = set(raw.get("documents_found") or [])
    missing = []
    if not any(d in found for d in ("BOM by Structure XML", "BOM by Structure PDF", "BOM Summary PDF")):
        missing.append("BOM")
    if "NCF" not in found:
        missing.append("NCF")
    if "Shop Drawing PDF" not in found:
        missing.append("Shop Drawing PDF")

    return {
        "job_code":             raw["job_code"],
        "project_name":         project_name,
        "project_name_source":  project_name_source,
        "year_released":        year_released,
        "date_released":        release_date,
        "sharepoint_year":      raw.get("sharepoint_year") or "",

        "location_bom":         loc_bom,
        "location_ncf":         loc_ncf,
        "location_conflict":    loc_conflict,
        "location_conflict_detail": loc_conflict_detail,
        "location":             location,
        "location_source":      location_source,
        "shipping_city":        city,
        "shipping_state":       state,
        "shipping_zip":         ncf.get("zipcode") or "",

        "contractor_bom":             contractor_bom,
        "customer_ncf":               customer_ncf,
        "contractor_customer_match":  contractor_customer_match,

        "structure_count_bom":             sc_bom,
        "structure_count_shop_drawing":    sc_shop,
        "structure_count_conflict":        sc_conflict,
        "structure_count_conflict_detail": sc_conflict_detail,

        "documents_found":   raw.get("documents_found") or [],
        "documents_missing": missing,
        "bom_file":          raw.get("bom_file") or "",
        "bom_doc_type":      raw.get("bom_doc_type") or "",
        "ncf_file":          raw.get("ncf_file") or "",
        "shop_drawing_file": raw.get("shop_drawing_file") or "",
        "parse_errors":      raw.get("parse_errors") or [],

        "in_sharepoint": True,
    }


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print("Authenticating with Microsoft Graph...")
    print("(If the token is expired you will be prompted to open a URL in your browser.)\n")
    token = acquire_token()
    print("Authenticated.\n")

    all_records = {}

    for year in ("2026", "2025", "2024", "2023"):
        cfg = YEAR_CONFIG[year]
        use_ncf = year != "2023"
        print(f"{'='*60}")
        print(f"[{year}]  {cfg['site_path']}  drive='{cfg['drive_name']}'")

        try:
            site  = get_site(token, HOSTNAME, cfg["site_path"])
            drive = get_drive(token, site["id"], cfg["drive_name"])
        except Exception as e:
            print(f"  ERROR connecting to {year} site: {e}")
            continue

        drive_id = drive["id"]
        top = list_children(token, drive_id, "root")
        job_folders = []
        for item in top:
            if "folder" not in item:
                continue
            m = _JOB_FOLDER_RE.match(item["name"].strip())
            job_code = m.group(1).upper() if m else item["name"].strip().upper()
            job_folders.append((job_code, item))

        print(f"  {len(job_folders)} job folders found\n")

        for i, (job_code, folder_item) in enumerate(job_folders, 1):
            token = ensure_fresh_token() or token
            print(f"  [{year}] {i}/{len(job_folders)}  {job_code} ...", end=" ", flush=True)
            try:
                raw    = process_job(token, drive_id, year, job_code, folder_item["id"], use_ncf)
                record = build_record(raw)

                if raw["parse_errors"]:
                    print(f"WARN  {raw['parse_errors'][0][:60]}", flush=True)
                else:
                    docs = ",".join(raw["documents_found"]) or "none"
                    print(f"OK  [{docs}]", flush=True)

                existing = all_records.get(job_code)
                if not existing:
                    all_records[job_code] = record
                elif record.get("location") and not existing.get("location"):
                    all_records[job_code] = record
                # Newest year processed first — existing record keeps priority if it has data

            except KeyboardInterrupt:
                print("\nInterrupted — saving partial results...")
                break
            except Exception as e:
                print(f"ERROR  {type(e).__name__}: {e}", flush=True)

        print()

    records = sorted(all_records.values(), key=lambda r: r["job_code"])

    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2, ensure_ascii=False)

    print(f"\nWrote {len(records):,} records -> {OUT_PATH}")

    n_loc       = sum(1 for r in records if r.get("location_conflict"))
    n_sc        = sum(1 for r in records if r.get("structure_count_conflict"))
    n_ct        = sum(1 for r in records if r.get("contractor_customer_match") is False)
    n_no_bom    = sum(1 for r in records if "BOM" in r.get("documents_missing", []))
    n_no_ncf    = sum(1 for r in records if "NCF" in r.get("documents_missing", []))
    n_errors    = sum(1 for r in records if r.get("parse_errors"))
    print(f"  Location conflicts (BOM vs NCF):         {n_loc:,}")
    print(f"  Structure count conflicts (BOM vs SD):   {n_sc:,}")
    print(f"  Contractor/customer mismatches:          {n_ct:,}")
    print(f"  Jobs missing BOM:                        {n_no_bom:,}")
    print(f"  Jobs missing NCF:                        {n_no_ncf:,}")
    print(f"  Jobs with parse errors:                  {n_errors:,}")


if __name__ == "__main__":
    main()
