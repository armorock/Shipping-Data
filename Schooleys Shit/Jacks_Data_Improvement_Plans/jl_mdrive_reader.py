"""
Walks the local M: Drive (2016-2022) and SharePoint Job Data sites (2023-2026),
parses each job folder's documents (BOM XML/PDF, NCF/ECF docx, Shop Drawing PDF),
and produces:

  output/jobcode_db_mdrive.json
    One record per job code with cross-document comparison fields identical
    to jobcode_db_sharepoint.json. Intended as an 8th source for jl_build_jobcode_db.py.

Year priority: oldest year is processed first; each successive year overwrites,
so 2026 naturally wins over 2016 when a job code appears in multiple years.

NCF parsing: enabled for SharePoint 2024-2026; disabled for 2023 and all M: Drive years.
Release folder: required for M: Drive jobs — folders without a "release" subfolder are skipped.

Run: python jl_mdrive_reader.py
"""

import io
import os
import re
import sys
import json
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

_BASE = os.path.dirname(os.path.abspath(__file__))
for _name in ("BOM-Structure-Detail", "BOM Structure Detail"):
    _p = os.path.abspath(os.path.join(_BASE, "..", "..", _name))
    if os.path.isdir(_p):
        sys.path.insert(0, _p)
        break

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

OUTPUT_DIR = os.path.join(_BASE, "output")
OUT_PATH   = os.path.join(OUTPUT_DIR, "jobcode_db_mdrive.json")

MDRIVE_ROOTS = {str(y): Path(f"M:\\{y}") for y in range(2016, 2023)}

HOSTNAME = "armorockllc.sharepoint.com"
SP_YEAR_CONFIG = {
    "2023": {"site_path": "/sites/jobdata2023", "drive_name": "Job Data 2023"},
    "2024": {"site_path": "/sites/jobdata2024", "drive_name": "Job Data 2024"},
    "2025": {"site_path": "/sites/jobdata2025", "drive_name": "Job Data 2025"},
    "2026": {"site_path": "/sites/JobData2026", "drive_name": "Job Data 2026"},
}

SKIP_FOLDERS = {"forms", "plugin_data", "robotinterface", "__macosx", "antiquated files"}

SP_MAX_WORKERS     = 12  # SharePoint Graph API calls — I/O-bound, safe up to ~20
MDRIVE_MAX_WORKERS =  6  # M: Drive SMB reads — lower to avoid hammering the NAS

_token_lock = threading.Lock()

_JOB_FOLDER_RE = re.compile(r"^([A-Ea-e][A-Za-z0-9]{2})\s*[-–—]\s*(.+)$")

_DOC_KEYWORDS = [
    ("bom by structure", ".xml",  "BOM by Structure XML"),
    ("bom by structure", ".pdf",  "BOM by Structure PDF"),
    ("bom summary",      ".pdf",  "BOM Summary PDF"),
    ("shop drawing",     ".pdf",  "Shop Drawing PDF"),
    ("shop drawings",    ".pdf",  "Shop Drawing PDF"),
    ("quote breakdown",  ".pdf",  "Quotation Breakdown PDF"),
    ("quote",            ".pdf",  "Quotation PDF"),
]
_DOC_LABEL_ORDER = {
    "BOM by Structure XML":    0,
    "BOM by Structure PDF":    1,
    "BOM Summary PDF":         2,
    "Shop Drawing PDF":        3,
    "Quotation Breakdown PDF": 4,
    "Quotation PDF":           5,
}

_PO_RE = re.compile(r"^[A-Z]{2,5}\s*[-–—]\s*(.+?)_PO#", re.IGNORECASE)


# ── Shared helpers ────────────────────────────────────────────────────────────

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


def _parse_city_state(loc_str):
    if not loc_str:
        return "", ""
    parts = loc_str.rsplit(",", 1)
    if len(parts) == 2:
        return parts[0].strip(), parts[1].strip()
    return loc_str.strip(), ""


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


def _empty_raw(job_code, year):
    return {
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
        "contractor_po_file":   None,
        "documents_found":      [],
        "parse_errors":         [],
    }


def build_record(raw):
    """Flatten raw parsed data into the canonical job record schema."""
    bom = raw.get("bom_header") or {}
    ncf = raw.get("ncf_meta") or {}

    project_name        = bom.get("job_name") or ncf.get("job_name") or ""
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

    loc_conflict        = bool(loc_bom and loc_ncf and loc_bom != loc_ncf)
    loc_conflict_detail = f"BOM='{loc_bom}'; NCF='{loc_ncf}'" if loc_conflict else None

    location        = loc_bom or loc_ncf or ""
    location_source = "BOM" if loc_bom else ("NCF" if loc_ncf else "")
    city, state     = _parse_city_state(location)

    contractor_bom = bom.get("contractor") or ""
    customer_ncf   = ncf.get("customer") or ""

    def _norm(s):
        return re.sub(r"\s+", " ", (s or "").strip().lower())

    contractor_customer_match = None
    if contractor_bom and customer_ncf:
        contractor_customer_match = _norm(contractor_bom) == _norm(customer_ncf)

    sc_bom  = raw.get("bom_structure_count")
    sc_shop = raw.get("shop_structure_count")
    sc_conflict        = bool(sc_bom is not None and sc_shop is not None and sc_bom != sc_shop)
    sc_conflict_detail = f"BOM={sc_bom}; ShopDrawing={sc_shop}" if sc_conflict else None

    found   = set(raw.get("documents_found") or [])
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

        "location_bom":             loc_bom,
        "location_ncf":             loc_ncf,
        "location_conflict":        loc_conflict,
        "location_conflict_detail": loc_conflict_detail,
        "location":                 location,
        "location_source":          location_source,
        "shipping_city":            city,
        "shipping_state":           state,
        "shipping_zip":             ncf.get("zipcode") or "",

        "contractor_bom":            contractor_bom,
        "contractor_po_file":        raw.get("contractor_po_file") or "",
        "customer_ncf":              customer_ncf,
        "contractor_customer_match": contractor_customer_match,

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


# ── M: Drive helpers ──────────────────────────────────────────────────────────

def _find_release_folder(job_path):
    try:
        for entry in os.scandir(job_path):
            if entry.is_dir() and "release" in entry.name.lower():
                return Path(entry.path)
    except PermissionError:
        pass
    return None


def _iter_release_files(folder, rel=""):
    try:
        entries = list(os.scandir(folder))
    except PermissionError:
        return
    for entry in entries:
        if entry.is_dir():
            if entry.name.lower() not in SKIP_FOLDERS:
                child_rel = f"{rel}/{entry.name}" if rel else entry.name
                yield from _iter_release_files(entry.path, child_rel)
        elif entry.is_file():
            yield Path(entry.path), rel


# ── Quote PDF helpers (copied from extract_mdrive.py) ────────────────────────

_PRICE_RE     = re.compile(r"\s*\$[\d,]+(?:\.\d{2})?$")
_PART_LINE_RE = re.compile(r"^(\d+)\s+(.+)")


def _dedouble_str(s):
    if len(s) >= 4 and len(s) % 2 == 0 and all(s[i] == s[i + 1] for i in range(0, len(s), 2)):
        return s[::2]
    return s


def _dedouble_words(text):
    return " ".join(_dedouble_str(w) for w in text.split())


def _infer_category(part_number, description):
    pn   = (part_number  or "").upper().strip()
    desc = (description  or "").upper()
    if pn.startswith("PR"):
        return "Precast"
    for pfx in ("MH", "RMH", "BOX", "SPECIAL"):
        if pn.startswith(pfx):
            return "Precast"
    if "PSX" in pn or "PSX" in desc:
        return "Connectors"
    if "MASTIC" in desc or "GASKET" in desc or pn.startswith("JM"):
        return "Joint Seal"
    if "RING AND COVER" in desc or "FRAME AND COVER" in desc or re.match(r"^A-\d", pn):
        return "Frame & Ring"
    if re.match(r"^\d{3}[A-Z]", pn) or re.match(r"^\d{2}[A-Z]{2}", pn):
        return "Precast"
    return "Miscellaneous"


def _split_part_tokens(token_str):
    token_str = _PRICE_RE.sub("", token_str).strip()
    parts = token_str.split(None, 1)
    if len(parts) == 1:
        return ("", parts[0]) if re.match(r"^\d", parts[0]) else (parts[0], "")
    pn_candidate, rest = parts
    if re.match(r"^[A-Za-z0-9][A-Za-z0-9._:-]*$", pn_candidate):
        return pn_candidate, rest.strip()
    return "", token_str.strip()


def _extract_quote_header(pages):
    header = {}
    for pg in pages[:2]:
        text = pg.extract_text() or ""
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            dl = _dedouble_words(line)
            if not header.get("job_name"):
                m = re.match(r"Ref(?:erence)?[:\s]+(.+)", dl, re.IGNORECASE)
                if m:
                    header["job_name"] = m.group(1).strip()
            if not header.get("release_date"):
                m = re.match(r"Quote\s+Date[:\s]+(\d{1,2}/\d{1,2}/\d{4})", dl, re.IGNORECASE)
                if m:
                    header["release_date"] = m.group(1)
            if not header.get("location"):
                if re.match(r"^[A-Za-z][A-Za-z ]+,\s*[A-Z]{2}$", dl):
                    header["location"] = dl
    return header


def parse_quote_breakdown_pdf(file_bytes):
    import pdfplumber
    structures = []
    current    = None
    pending    = None

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        header = _extract_quote_header(pdf.pages)
        for pg in pdf.pages:
            text = pg.extract_text() or ""
            for raw_line in text.splitlines():
                line = raw_line.strip()
                if not line:
                    continue
                m = re.match(r"Structure\s+ID\s*:\s*(.+?)(?:\s+\d+\s+of\s+\d+)?\s*$", line, re.IGNORECASE)
                if m:
                    current = {"structure_name": m.group(1).strip(), "line_items": []}
                    structures.append(current)
                    pending = None
                    continue
                if re.match(r"^Qty\b|^Total\b|^For:|^Ref:|^Quote|^Page\s+\d", line, re.IGNORECASE):
                    pending = None
                    continue
                if current is None:
                    continue
                pm = _PART_LINE_RE.match(line)
                if pm:
                    qty  = pm.group(1)
                    pn, desc = _split_part_tokens(pm.group(2))
                    item = {"part_number": pn, "description": desc, "quantity": qty,
                            "weight_lbs": "", "category": _infer_category(pn, desc)}
                    current["line_items"].append(item)
                    pending = item
                elif pending and not re.search(r"\$[\d,]+", line):
                    continuation = line.strip("()")
                    pending["description"] = (pending["description"] + " " + continuation).strip()

    return {"header": header, "structures": structures}


def parse_quote_pdf(file_bytes):
    import pdfplumber
    structures = []
    current    = None
    PRICE_COL  = 410
    RIGHT_COL  = 310

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        header = _extract_quote_header(pdf.pages)
        for pg in pdf.pages:
            words = pg.extract_words()
            if not words:
                continue
            lines_by_y = {}
            for w in words:
                y = round(w["top"])
                lines_by_y.setdefault(y, []).append(w)
            for y in sorted(lines_by_y):
                wds   = sorted(lines_by_y[y], key=lambda w: w["x0"])
                texts = [w["text"] for w in wds]
                if any(t.startswith("$") for t in texts):
                    name_words = [_dedouble_str(w["text"]) for w in wds if w["x0"] < PRICE_COL]
                    structure_name = " ".join(name_words).strip()
                    current = {"structure_name": structure_name, "line_items": []}
                    structures.append(current)
                    continue
                if current is None:
                    continue
                if texts and texts[0].lower() in ("structure", "for:", "ref:", "qty", "description"):
                    continue
                for col_lo, col_hi in [(70, RIGHT_COL), (RIGHT_COL, 620)]:
                    col_wds = [w for w in wds if col_lo <= w["x0"] < col_hi]
                    if not col_wds or not col_wds[0]["text"].isdigit():
                        continue
                    qty      = col_wds[0]["text"]
                    rest_str = " ".join(w["text"] for w in col_wds[1:])
                    pn, desc = _split_part_tokens(rest_str)
                    if current is not None:
                        current["line_items"].append({
                            "part_number": pn, "description": desc, "quantity": qty,
                            "weight_lbs": "", "category": _infer_category(pn, desc),
                        })

    return {"header": header, "structures": structures}


def process_mdrive_job(job_path, job_code, year):
    raw = _empty_raw(job_code, year)

    release_folder = _find_release_folder(job_path)
    scan_root = release_folder if release_folder is not None else job_path

    # PO files live in the job root, not necessarily in the release subfolder
    try:
        for entry in os.scandir(job_path):
            if entry.is_file():
                m = _PO_RE.match(entry.name)
                if m:
                    raw["contractor_po_file"] = m.group(1).strip()
                    break
    except PermissionError:
        pass

    all_files = list(_iter_release_files(scan_root))

    bom_files = [(p, rel) for p, rel in all_files if _classify_doc(p.name)]
    bom_files.sort(key=lambda t: (
        _DOC_LABEL_ORDER.get(_classify_doc(t[0].name), 99),
        tuple(-x for x in _filename_sort_date(t[0].name)),
    ))

    for fpath, _ in bom_files:
        fname = fpath.name
        label = _classify_doc(fname)
        try:
            data = fpath.read_bytes()

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
                if raw["shop_structure_count"] is None:
                    raw["shop_structure_count"] = len(parsed.get("structures", []))
                    raw["shop_drawing_file"]    = fname
                    raw["documents_found"].append("Shop Drawing PDF")
            else:
                if raw["bom_header"] is None:
                    raw["bom_header"]          = parsed.get("header", {})
                    raw["bom_structure_count"] = len(parsed.get("structures", [])) if "structures" in parsed else None
                    raw["bom_file"]            = fname
                    raw["bom_doc_type"]        = label
                    raw["documents_found"].append(label)

        except Exception as e:
            raw["parse_errors"].append(f"{label}({fname}): {e}")

    if raw["bom_header"] is None:
        quote_files = [
            (p, rel) for p, rel in all_files
            if _classify_doc(p.name) in ("Quotation Breakdown PDF", "Quotation PDF")
        ]
        quote_files.sort(key=lambda t: _DOC_LABEL_ORDER.get(_classify_doc(t[0].name), 99))
        for fpath, _ in quote_files:
            fname = fpath.name
            label = _classify_doc(fname)
            try:
                data = fpath.read_bytes()
                parsed = parse_quote_breakdown_pdf(data) if label == "Quotation Breakdown PDF" else parse_quote_pdf(data)
                if not parsed:
                    continue
                hdr     = parsed.get("header", {})
                structs = parsed.get("structures", [])
                raw["bom_header"]          = hdr
                raw["bom_structure_count"] = len(structs)
                raw["bom_file"]            = fname
                raw["bom_doc_type"]        = label
                raw["documents_found"].append(label)
                break
            except Exception as e:
                raw["parse_errors"].append(f"{label}({fname}): {e}")

    return raw


# ── SharePoint helpers ────────────────────────────────────────────────────────

def _iter_sp_folder(token, drive_id, folder_id, rel_path=""):
    for item in list_children(token, drive_id, folder_id):
        if "file" in item:
            yield item, rel_path
        elif "folder" in item:
            name = item["name"].lower()
            if name not in SKIP_FOLDERS:
                child = f"{rel_path}/{item['name']}" if rel_path else item["name"]
                yield from _iter_sp_folder(token, drive_id, item["id"], child)


def process_sp_job(token, drive_id, year, job_code, folder_id, use_ncf=True):
    raw      = _empty_raw(job_code, year)
    all_files = list(_iter_sp_folder(token, drive_id, folder_id))

    bom_files      = [(f, p) for f, p in all_files if _classify_doc(f["name"])]
    ncf_candidates = [f for f, _ in all_files if _is_ncf(f["name"])]
    if not ncf_candidates and use_ncf:
        ncf_candidates = [f for f, _ in all_files if f["name"].lower().endswith(".docx")]

    bom_files.sort(key=lambda t: (
        _DOC_LABEL_ORDER.get(_classify_doc(t[0]["name"]), 99),
        tuple(-x for x in _filename_sort_date(t[0]["name"])),
    ))

    for file_item, _ in all_files:
        m = _PO_RE.match(file_item["name"])
        if m:
            raw["contractor_po_file"] = m.group(1).strip()
            break

    if use_ncf:
        for ncf_item in ncf_candidates:
            try:
                data   = download_file(token, drive_id, ncf_item["id"])
                parsed = _safe_ncf(data)
                if parsed:
                    if raw["ncf_meta"] is None or parsed.get("location"):
                        raw["ncf_meta"] = parsed
                        raw["ncf_file"] = ncf_item["name"]
                    if parsed.get("location"):
                        break
            except Exception as e:
                raw["parse_errors"].append(f"NCF({ncf_item['name']}): {e}")
        if raw["ncf_meta"] is not None:
            raw["documents_found"].append("NCF")

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
                if raw["shop_structure_count"] is None:
                    raw["shop_structure_count"] = len(parsed.get("structures", []))
                    raw["shop_drawing_file"]    = fname
                    raw["documents_found"].append("Shop Drawing PDF")
            else:
                if raw["bom_header"] is None:
                    raw["bom_header"]          = parsed.get("header", {})
                    raw["bom_structure_count"] = len(parsed.get("structures", [])) if "structures" in parsed else None
                    raw["bom_file"]            = fname
                    raw["bom_doc_type"]        = label
                    raw["documents_found"].append(label)

        except Exception as e:
            raw["parse_errors"].append(f"{label}({fname}): {e}")

    return raw


# ── Year runners ──────────────────────────────────────────────────────────────

def run_mdrive(all_records):
    for year in sorted(MDRIVE_ROOTS):
        root = MDRIVE_ROOTS[year]
        print(f"{'='*60}")
        print(f"[M: Drive {year}]  {root}")

        if not root.exists():
            print(f"  {root} not found — skipping")
            continue

        try:
            job_folders = sorted(p for p in root.iterdir() if p.is_dir())
        except PermissionError as e:
            print(f"  PermissionError reading {root}: {e}")
            continue

        n = len(job_folders)
        print(f"  {n} folders")

        def _mdrive_worker(job_path):
            fname = job_path.name.strip()
            m = _JOB_FOLDER_RE.match(fname)
            if not m:
                return None
            jc = m.group(1).upper()
            raw    = process_mdrive_job(job_path, jc, year)
            record = build_record(raw)
            return jc, record, raw

        done = 0
        with ThreadPoolExecutor(max_workers=MDRIVE_MAX_WORKERS) as pool:
            futures = {pool.submit(_mdrive_worker, jp): jp.name for jp in job_folders}
            try:
                for fut in as_completed(futures):
                    result = fut.result()
                    if result is None:
                        continue
                    done += 1
                    job_code, record, raw = result
                    docs = ",".join(raw["documents_found"]) or "none"
                    if raw["parse_errors"]:
                        print(f"  [{year}] {done}/{n}  {job_code}  WARN  {raw['parse_errors'][0][:60]}", flush=True)
                    else:
                        print(f"  [{year}] {done}/{n}  {job_code}  [{docs}]", flush=True)
                    all_records[job_code] = record
            except KeyboardInterrupt:
                print("\nInterrupted.")
                raise

        print()


def run_sharepoint(all_records, token):
    for year in sorted(SP_YEAR_CONFIG):
        cfg     = SP_YEAR_CONFIG[year]
        use_ncf = year != "2023"
        print(f"{'='*60}")
        print(f"[SharePoint {year}]  {cfg['site_path']}  drive='{cfg['drive_name']}'")

        try:
            site  = get_site(token, HOSTNAME, cfg["site_path"])
            drive = get_drive(token, site["id"], cfg["drive_name"])
        except Exception as e:
            print(f"  ERROR connecting to {year} site: {e}")
            continue

        drive_id    = drive["id"]
        top         = list_children(token, drive_id, "root")
        job_folders = []
        for item in top:
            if "folder" not in item:
                continue
            fm = _JOB_FOLDER_RE.match(item["name"].strip())
            job_code = fm.group(1).upper() if fm else item["name"].strip().upper()
            job_folders.append((job_code, item))

        n = len(job_folders)
        print(f"  {n} job folders\n")

        token_box = [token]

        def _sp_worker(args):
            job_code, folder_item = args
            with _token_lock:
                token_box[0] = ensure_fresh_token() or token_box[0]
                t = token_box[0]
            try:
                raw    = process_sp_job(t, drive_id, year, job_code, folder_item["id"], use_ncf)
                record = build_record(raw)
                return job_code, record, raw, None
            except Exception as e:
                return job_code, None, {"documents_found": [], "parse_errors": []}, e

        done = 0
        with ThreadPoolExecutor(max_workers=SP_MAX_WORKERS) as pool:
            futures = {pool.submit(_sp_worker, item): item[0] for item in job_folders}
            try:
                for fut in as_completed(futures):
                    done += 1
                    job_code, record, raw, err = fut.result()
                    if err:
                        print(f"  [{year}] {done}/{n}  {job_code}  ERROR  {type(err).__name__}: {err}", flush=True)
                    elif raw["parse_errors"]:
                        print(f"  [{year}] {done}/{n}  {job_code}  WARN  {raw['parse_errors'][0][:60]}", flush=True)
                    else:
                        docs = ",".join(raw["documents_found"]) or "none"
                        print(f"  [{year}] {done}/{n}  {job_code}  OK  [{docs}]", flush=True)
                    if record:
                        all_records[job_code] = record
            except KeyboardInterrupt:
                print("\nInterrupted — saving partial results...")
                raise

        token = token_box[0]
        print()


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    all_records = {}

    run_mdrive(all_records)

    print("Authenticating with Microsoft Graph...")
    print("(If the token is expired you will be prompted to open a URL in your browser.)\n")
    token = acquire_token()
    print("Authenticated.\n")

    run_sharepoint(all_records, token)

    records = sorted(all_records.values(), key=lambda r: r["job_code"])

    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2, ensure_ascii=False)

    print(f"\nWrote {len(records):,} records -> {OUT_PATH}")

    n_loc    = sum(1 for r in records if r.get("location_conflict"))
    n_sc     = sum(1 for r in records if r.get("structure_count_conflict"))
    n_ct     = sum(1 for r in records if r.get("contractor_customer_match") is False)
    n_no_bom = sum(1 for r in records if "BOM" in r.get("documents_missing", []))
    n_no_ncf = sum(1 for r in records if "NCF" in r.get("documents_missing", []))
    n_errors = sum(1 for r in records if r.get("parse_errors"))
    print(f"  Total records:                           {len(records):,}")
    print(f"  Location conflicts (BOM vs NCF):         {n_loc:,}")
    print(f"  Structure count conflicts (BOM vs SD):   {n_sc:,}")
    print(f"  Contractor/customer mismatches:          {n_ct:,}")
    print(f"  Jobs missing BOM:                        {n_no_bom:,}")
    print(f"  Jobs missing NCF:                        {n_no_ncf:,}")
    print(f"  Jobs with parse errors:                  {n_errors:,}")


if __name__ == "__main__":
    main()
