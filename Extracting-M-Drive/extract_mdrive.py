import csv
import io
import os
import re
import sys
from datetime import date
from pathlib import Path

from parse_bom_pdf import (
    parse_bom_pdf_safe,
    parse_bom_by_structure_pdf_safe,
    parse_bom_by_structure_xml,
    parse_shop_drawing_pdf_safe,
    normalize_location,
)

YEAR_ROOTS = {str(y): Path(f"M:\\{y}") for y in range(2016, 2025)}
SKIP_FOLDERS = {"ar folder", "ar files", "antiquated", "__macosx"}
OUTPUT_DIR = "output"

_today = date.today()
DATE_EXTRACTED = f"{_today.month}/{_today.day}/{_today.year}"

_JOB_FOLDER_RE = re.compile(r"^([A-Ea-e][A-Za-z]{2})\s*[-–—]\s*(.+)$")

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
    ("SPECIAL", "Special"),
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
# ── Quote PDF helpers ─────────────────────────────────────────────────────────

_PRICE_RE    = re.compile(r"\s*\$[\d,]+(?:\.\d{2})?$")
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
        return "Precast"  # Pro-Ring; classify_part re-routes to Resale
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
        return "Precast"  # polymer products: 603C, 604B, 60EL, etc.
    return "Miscellaneous"


def _split_part_tokens(token_str):
    """Split 'PNUMBER description...' stripping trailing price."""
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
    """Parse two-column Armorock QUOTE PDFs."""
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
            for kw, v in [("CONE", "Cone"), ("LID", "Lid"), ("SECTION", "Section"),
                          ("BASE", "Base"), ("GRADE RING", "Grade Ring"), ("RISER", "Riser")]:
                if kw in desc_up:
                    part_type = v
                    break
        subtype = next((v for k, v in _PRECAST_SUBTYPE_MAP if k in desc_up), "")
        return part_type, subtype
    else:
        subtype = next((v for k, v in _RESALE_SUBTYPE_MAP if k in desc_up), "")
        return "Resale", subtype


# ── Skip logging ─────────────────────────────────────────────────────────────

def _write_skipped(job_code, filename, structure_name, year, skipped_path):
    with open(skipped_path, "a", newline="", encoding="utf-8") as f:
        csv.DictWriter(f, fieldnames=["year", "job_code", "structure_name", "source_file"]).writerow(
            {"year": year, "job_code": job_code, "structure_name": structure_name, "source_file": filename}
        )


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


# ── File type detection (table-driven) ───────────────────────────────────────

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


def classify_document(filename):
    n = filename.lower()
    ext = os.path.splitext(n)[1]
    for keyword, extension, label in _DOC_TYPES:
        if keyword in n and ext == extension:
            return label
    return None


# ── Local file traversal ──────────────────────────────────────────────────────

def _find_release_folder(job_path):
    try:
        for entry in os.scandir(job_path):
            if entry.is_dir() and "release" in entry.name.lower():
                return Path(entry.path)
    except PermissionError:
        pass
    return None


def _iter_bom_files(folder, rel=""):
    try:
        entries = list(os.scandir(folder))
    except PermissionError:
        return
    for entry in entries:
        if entry.is_dir():
            if entry.name.lower() not in SKIP_FOLDERS:
                child_rel = f"{rel}/{entry.name}" if rel else entry.name
                yield from _iter_bom_files(entry.path, child_rel)
        elif entry.is_file() and classify_document(entry.name):
            yield Path(entry.path), rel


# ── Row builder ───────────────────────────────────────────────────────────────

def _build_row(job_code, drive_year, header, structure_name, item,
               source_file, source_subfolder, bom_type):
    part_type, subtype = classify_part(item["category"], item["part_number"], item["description"])
    pn_name = build_part_name(item["part_number"])
    if pn_name == (item.get("part_number") or "").strip():
        pn_name = item.get("description") or pn_name

    bom_location = header.get("location", "")
    job_location = normalize_location(bom_location) if bom_location else ""
    location_source = "BOM" if job_location else ""
    if not job_location and job_code in _LOCATION_OVERRIDES:
        ov = _LOCATION_OVERRIDES[job_code]
        job_location    = ov["location"]
        location_source = ov["source"]

    release_date = header.get("release_date", "")
    if release_date:
        parts = release_date.split("/")
        year_val_out = parts[2] if len(parts) == 3 and parts[2].isdigit() else drive_year
    else:
        year_val_out = drive_year

    return {
        "Year Release":     year_val_out,
        "BOM Release Date": release_date,
        "Date extracted":   DATE_EXTRACTED,
        "Job Code":         job_code,
        "Project Name":     header.get("job_name", ""),
        "Structure Name":   structure_name,
        "Job Location":     job_location,
        "Location Source":  location_source,
        "Zip Code":         "",
        "Contractor":       header.get("contractor", ""),
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
    base = re.sub(r"-Selected-BOM Summary\.pdf$", "", filename, flags=re.IGNORECASE)
    base = re.sub(r"-BOM Summary\.pdf$", "", base, flags=re.IGNORECASE)
    clean_job = (pdf_job_name or "").strip()
    if clean_job and base.startswith(clean_job):
        rest = base[len(clean_job):].lstrip("- ").strip()
    else:
        rest = base.strip()
    return rest or pdf_job_name or ""


# ── Per-job processor ─────────────────────────────────────────────────────────

def process_job(job_path, job_code, drive_year, skipped_path):
    release_folder = _find_release_folder(job_path)
    if release_folder is None:
        return []

    bom_files = list(_iter_bom_files(release_folder))
    if not bom_files:
        return []

    if not re.match(r"^[A-Ea-e][A-Za-z]{2}$", job_code):
        for fpath, _ in bom_files:
            m = re.match(r"^([A-Ea-e][A-Za-z]{2})[\s-]", fpath.name)
            if m:
                job_code = m.group(1).upper()
                break

    rows = []
    for fpath, subfolder in bom_files:
        fname = fpath.name
        label = classify_document(fname)
        print(f"  [{job_code}] ({label}) {subfolder or '.'}/{fname}")
        try:
            file_bytes = fpath.read_bytes()
            if label == "BOM by Structure XML":
                bom = parse_bom_by_structure_xml(file_bytes)
                for skipped in bom.get("skipped_structures", []):
                    print(f'  [{job_code}] [SKIP structure] "{skipped}" in {fname} — no pattern matched')
                    _write_skipped(job_code, fname, skipped, drive_year, skipped_path)
                for structure in bom["structures"]:
                    sname = structure["structure_name"]
                    for item in structure["line_items"]:
                        rows.append(_build_row(job_code, drive_year, bom["header"], sname,
                                               item, fname, subfolder, label))
            elif label == "BOM by Structure PDF":
                bom = parse_bom_by_structure_pdf_safe(file_bytes)
                for skipped in bom.get("skipped_structures", []):
                    print(f'  [{job_code}] [SKIP structure] "{skipped}" in {fname} — no pattern matched')
                    _write_skipped(job_code, fname, skipped, drive_year, skipped_path)
                for structure in bom["structures"]:
                    sname = structure["structure_name"]
                    for item in structure["line_items"]:
                        rows.append(_build_row(job_code, drive_year, bom["header"], sname,
                                               item, fname, subfolder, label))
            elif label == "Shop Drawing PDF":
                bom = parse_shop_drawing_pdf_safe(file_bytes)
                for skipped in bom.get("skipped_structures", []):
                    print(f'  [{job_code}] [SKIP structure] "{skipped}" in {fname} — no pattern matched')
                    _write_skipped(job_code, fname, skipped, drive_year, skipped_path)
                for structure in bom["structures"]:
                    sname = structure["structure_name"]
                    for item in structure["line_items"]:
                        rows.append(_build_row(job_code, drive_year, bom["header"], sname,
                                               item, fname, subfolder, label))
            elif label == "Quotation PDF":
                fn_lower = fname.lower()
                parser = parse_quote_breakdown_pdf if ("quote breakdown" in fn_lower or "quotation detail" in fn_lower) else parse_quote_pdf
                bom = parser(file_bytes)
                for structure in bom["structures"]:
                    sname = structure["structure_name"]
                    for item in structure["line_items"]:
                        rows.append(_build_row(job_code, drive_year, bom["header"], sname,
                                               item, fname, subfolder, label))
            elif label == "BOM Summary PDF":
                bom = parse_bom_pdf_safe(file_bytes)
                sname = _parse_summary_structure(fname, bom["header"].get("job_name", ""))
                for item in bom["line_items"]:
                    rows.append(_build_row(job_code, drive_year, bom["header"], sname,
                                           item, fname, subfolder, label))
        except BaseException as exc:
            print(f"  [{job_code}] ERROR parsing {fname}: {type(exc).__name__}: {exc}")
            if isinstance(exc, KeyboardInterrupt):
                raise
    return rows


# ── Year runner ───────────────────────────────────────────────────────────────

def run_year(year):
    year_root = YEAR_ROOTS[year]
    output_path  = os.path.join(OUTPUT_DIR, f"all_bom_{year}.csv")
    skipped_path = os.path.join(OUTPUT_DIR, f"skipped_structures_{year}.csv")

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(skipped_path, "w", newline="", encoding="utf-8") as f:
        csv.DictWriter(f, fieldnames=["year", "job_code", "structure_name", "source_file"]).writeheader()

    print(f"\n{'='*60}")
    print(f"Year: {year}  root={year_root}")

    if not year_root.exists():
        print(f"  {year_root} not found — skipping")
        return []

    job_folders = sorted(p for p in year_root.iterdir() if p.is_dir())
    print(f"{len(job_folders)} job folders")

    all_rows = []
    errors   = []

    for job_path in job_folders:
        fname = job_path.name.strip()
        m = _JOB_FOLDER_RE.match(fname)
        job_code = m.group(1) if m else re.split(r"[-\s]", fname)[0].strip()
        try:
            rows = process_job(job_path, job_code, year, skipped_path)
            all_rows.extend(rows)
        except KeyboardInterrupt:
            raise
        except Exception as exc:
            msg = f"[{job_code}] FOLDER ERROR: {type(exc).__name__}: {exc}"
            print(f"  {msg}")
            errors.append(msg)

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
    valid = list(YEAR_ROOTS.keys())
    years = sys.argv[1:] if len(sys.argv) > 1 else valid
    for y in years:
        if y not in YEAR_ROOTS:
            print(f"Unknown year: {y}  (valid: {valid})")
            sys.exit(1)
    for y in years:
        run_year(y)


if __name__ == "__main__":
    main()
