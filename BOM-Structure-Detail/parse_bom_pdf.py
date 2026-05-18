"""
Parsers for Armorock BOM documents.

parse_bom_pdf(pdf_bytes)
  → {header, opening_schedule, line_items, total_precast_weight}
  For "BOM Summary" PDFs (single page per job, all structures combined).

parse_bom_by_structure_pdf(pdf_bytes)
  → {header, structures}
  For "BOM by Structure" PDFs (one page per structure/manhole).
  structures: [{structure_name, line_items, total_precast_weight}]

parse_bom_by_structure_xml(xml_bytes)
  → {header, structures}
  For "BOM by Structure (Excel XML).xml" SpreadsheetML files.
  Same output shape as parse_bom_by_structure_pdf.
"""

import re
import io
import os as _os
import csv as _csv_io

_HERE = _os.path.dirname(_os.path.abspath(__file__))

def _load_state_abbr():
    path = _os.path.join(_HERE, "data", "state_abbreviations.csv")
    if not _os.path.exists(path):
        return {}
    abbr = {}
    with open(path, newline="", encoding="utf-8") as f:
        for row in _csv_io.DictReader(f):
            abbr[row["Full Name"].lower()] = row["Abbreviation"]
    return abbr

_STATE_ABBR = _load_state_abbr()


def normalize_location(loc):
    if not loc:
        return loc
    parts = loc.rsplit(",", 1)
    if len(parts) == 2:
        city, state = parts[0].strip(), parts[1].strip()
        norm = _STATE_ABBR.get(state.lower(), state)
        return f"{city}, {norm}"
    return _STATE_ABBR.get(loc.lower().strip(), loc)


def parse_bom_pdf(pdf_bytes):
    try:
        import pdfplumber
    except ImportError:
        raise ImportError("pdfplumber is required: pip install pdfplumber")

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        page = pdf.pages[0]
        text = page.extract_text(layout=True) or ""
        tables = page.extract_tables()
        release_date = _parse_release_date(page)

    header = _parse_header(text)
    header["release_date"] = release_date
    return {
        "header": header,
        "opening_schedule": _parse_opening_schedule(tables, text),
        "line_items": _parse_line_items(tables, text),
        "total_precast_weight": _parse_total_weight(text),
    }


def parse_bom_by_structure_pdf(pdf_bytes):
    """Parse a BOM by Structure PDF (one table per manhole/structure).

    Returns {header, structures} where structures is a list of
    {structure_name, line_items, total_precast_weight}.
    """
    try:
        import pdfplumber
    except ImportError:
        raise ImportError("pdfplumber is required: pip install pdfplumber")

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        first_page = pdf.pages[0]
        first_text = first_page.extract_text(layout=True) or ""
        header = _parse_header(first_text)
        header["release_date"] = _parse_release_date(first_page)

        structures = []
        skipped_structures = []
        for page in pdf.pages:
            for table in (page.extract_tables() or []):
                if len(table) < 3 or not table[0] or not table[0][0]:
                    continue
                name_cell = str(table[0][0]).strip()
                name = re.split(r"\s+Ship Date", name_cell, flags=re.IGNORECASE)[0].strip()
                if any(p.search(name) for p in _NON_STRUCTURE_PATTERNS):
                    continue
                if not any(p.search(name) for p in _STRUCTURE_PATTERNS):
                    skipped_structures.append(name)
                    continue
                items_text = str(table[2][0] or "") if table[2] and table[2][0] else ""
                items = _line_items_from_text(items_text, start_in_items=True)
                total = _parse_total_weight(items_text)
                if items:
                    structures.append({
                        "structure_name":       name,
                        "line_items":           items,
                        "total_precast_weight": total,
                    })

    return {"header": header, "structures": structures, "skipped_structures": skipped_structures}


def parse_bom_by_structure_xml(xml_bytes):
    """Parse a BOM by Structure (Excel XML) SpreadsheetML file.

    Returns {header, structures} with the same shape as parse_bom_by_structure_pdf.
    Column layout: structure_name(0), internal_id(1), item_type(2),
                   part_number(3), part_number_alt(4), description(5),
                   quantity(6), weight(7).
    """
    import xml.etree.ElementTree as ET

    _NS = "urn:schemas-microsoft-com:office:spreadsheet"
    _ITEM_TYPE_MAP = {
        "StackElement": "Precast",
        "Connector":    "Connectors",
        "JointSeal":    "Joint Seal",
        "MiscItem":     "Miscellaneous",
        "Frame":        "Frame & Ring",
    }
    _SKIP_ITEM_TYPES = {"Hole"}

    def _row_cells(row):
        cells = [""] * 16
        idx = 0
        for cell in row.findall(f"{{{_NS}}}Cell"):
            attr_idx = cell.get(f"{{{_NS}}}Index")
            if attr_idx:
                idx = int(attr_idx) - 1
            if idx < len(cells):
                data = cell.find(f"{{{_NS}}}Data")
                cells[idx] = (data.text or "").strip() if data is not None else ""
            idx += 1
        return cells

    root = ET.fromstring(xml_bytes)
    ws = root.find(f".//{{{_NS}}}Worksheet")
    if ws is None:
        return {"header": _empty_header(), "structures": []}
    table = ws.find(f"{{{_NS}}}Table")
    if table is None:
        return {"header": _empty_header(), "structures": []}

    header = _empty_header()
    structures = {}
    structure_order = []

    for row in table.findall(f"{{{_NS}}}Row"):
        cells = _row_cells(row)
        item_type = cells[2]

        if item_type in _ITEM_TYPE_MAP or item_type in _SKIP_ITEM_TYPES:
            if item_type in _SKIP_ITEM_TYPES:
                continue
            structure_name = cells[0]
            if not structure_name or structure_name.lower().startswith("x"):
                continue
            if structure_name not in structures:
                structures[structure_name] = []
                structure_order.append(structure_name)
            structures[structure_name].append({
                "category":    _ITEM_TYPE_MAP[item_type],
                "quantity":    cells[6],
                "part_number": cells[3],
                "description": cells[5],
                "weight_lbs":  cells[7].replace(",", "") if cells[7] else "",
            })
        else:
            for i, val in enumerate(cells):
                if not val:
                    continue
                label = val.rstrip(":")
                nxt = next((cells[j] for j in range(i + 1, len(cells)) if cells[j]), "")
                if label == "Job Name" and not header["job_name"]:
                    header["job_name"] = nxt
                elif label == "Job Number" and not header["job_number"]:
                    header["job_number"] = nxt
                elif label == "Location" and not header["location"]:
                    header["location"] = nxt
                elif label in ("Contractor", "Contractor Name") and not header["contractor"]:
                    header["contractor"] = nxt
                elif label == "Reported On" and not header["release_date"]:
                    dt_m = re.match(r"(\d{4})-(\d{2})-(\d{2})", nxt)
                    if dt_m:
                        header["release_date"] = f"{int(dt_m.group(2))}/{int(dt_m.group(3))}/{dt_m.group(1)}"
                elif re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", val) and not header["release_date"]:
                    header["release_date"] = val

    result_structures = [
        {"structure_name": name, "line_items": structures[name], "total_precast_weight": None}
        for name in structure_order
        if structures[name]
    ]
    return {"header": header, "structures": result_structures}


def _empty_header():
    return {"job_name": "", "job_number": "", "location": "", "contractor": "", "phone": "", "release_date": ""}


# ---------------------------------------------------------------------------
# Structure name detection (BOM by Structure)
# ---------------------------------------------------------------------------

_NON_STRUCTURE_PATTERNS = [
    re.compile(r"^Bill of Material", re.IGNORECASE),
    re.compile(r"^[x_]", re.IGNORECASE),
    re.compile(r"^(?:Spreader|Lifting|Rigging)\b", re.IGNORECASE),
    re.compile(r"^(?:Wrapid|Grout\s+Kit|Joint\s+Seal|Riser\s+Wrap)", re.IGNORECASE),
    re.compile(r"^(?:Extra|Additional)\s+(?:Pro-|Rings|Risers)", re.IGNORECASE),
]

_STRUCTURE_PATTERNS = [
    re.compile(r"\bSTA\s+\d+\+\d+", re.IGNORECASE),
    re.compile(r"\bSSMH\b", re.IGNORECASE),
    re.compile(r"^[A-Z]{0,2}MH[-\s#\d]", re.IGNORECASE),
    re.compile(r"^ARV[M]?\s*[-\s\d]", re.IGNORECASE),
    re.compile(r"^(?:New|Proposed|Rehab|Receiving)\s+(?:MH|Manhole)", re.IGNORECASE),
    re.compile(r"^(?:Phase|PH)\s+\d", re.IGNORECASE),
    re.compile(r"^\d+\+\d+", re.IGNORECASE),
    re.compile(r"^\d+-\d+", re.IGNORECASE),
    re.compile(r"^[A-Z]{2}\d{2}-\d{2}", re.IGNORECASE),
    re.compile(r"^(?:MANHOLE|Manhole)\b", re.IGNORECASE),
    re.compile(r"^(?:INLET|JUNCTION|CATCH\s*BASIN|CB|VAULT)\s+", re.IGNORECASE),
    re.compile(r"^(?:LS|WW|PS)(?:[-\s#\d]|$)", re.IGNORECASE),
    re.compile(r"^[A-Z]{1,2}\d+[A-Z]?(\s*\(|\s*$)", re.IGNORECASE),
    re.compile(r"^(?:SMH|DMH|CMH)\b", re.IGNORECASE),
    re.compile(r"\bNO\.?\s*\d+", re.IGNORECASE),
    re.compile(r"^(?:STR|STRUCTURE)\s*[-\s]?[\d\w]", re.IGNORECASE),
    re.compile(r"^[A-Z]{1,3}-[\dA-Z]+$", re.IGNORECASE),
    re.compile(r"^TYPE\s+[\dA-Z]", re.IGNORECASE),
    re.compile(r"^(?:CO|CLEANOUT|ACCESS\s+POINT)\s*[-\s]?[\d\w]", re.IGNORECASE),
    re.compile(r"^(?:Wet\s+Well|Pump\s+Station|Lift\s+Station|Sump\b|Valve\s+Vault|Influent\b|Metering\s+Vault)", re.IGNORECASE),
]


def _find_structure_name(text):
    """Return the structure identifier from the first matching line of page text."""
    for line in text.split("\n"):
        stripped = line.strip()
        if not stripped:
            continue
        for pat in _STRUCTURE_PATTERNS:
            if pat.search(stripped):
                return re.split(r"\s+Ship Date", stripped, flags=re.IGNORECASE)[0].strip()
    return ""


# ---------------------------------------------------------------------------
# Release date (bottom-right corner of page)
# ---------------------------------------------------------------------------

def _parse_release_date(page):
    words = page.extract_words()
    h = float(page.height)
    w = float(page.width)
    candidates = [wd for wd in words if wd["top"] > h * 0.88 and wd["x0"] > w * 0.60]
    for wd in sorted(candidates, key=lambda x: -x["top"]):
        if re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", wd["text"]):
            return wd["text"]
    return ""


# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------

def _parse_header(text):
    def field(label):
        m = re.search(rf"{re.escape(label)}[^\S\n]*([^\n]*)", text)
        val = (m.group(1) or "").strip() if m else ""
        # Strip anything that looks like the next label bleeding onto the same line
        val = re.split(r"\s{3,}|\t", val)[0].strip()
        return val

    return {
        "job_name":   field("Job Name:"),
        "job_number": field("Job Number:"),
        "location":   field("Location:"),
        "contractor": field("Contractor:"),
        "phone":      field("Phone:"),
    }


# ---------------------------------------------------------------------------
# Opening schedule
# ---------------------------------------------------------------------------

def _parse_opening_schedule(tables, text):
    # Try to find it in a pdfplumber table first
    for table in (tables or []):
        rows = _opening_schedule_from_table(table)
        if rows:
            return rows

    # Fallback: parse from raw text
    return _opening_schedule_from_text(text)


def _opening_schedule_from_table(table):
    if not table:
        return []
    flat = [str(c).strip() for row in table for c in (row or []) if c]
    if not any("Pipe Size" in s or "Opening Schedule" in s for s in flat):
        return []

    rows = []
    header_seen = False
    for row in table:
        cols = [str(c).strip() if c else "" for c in row]
        if not header_seen:
            if any("Pipe Size" in c for c in cols):
                header_seen = True
            continue
        # Stop at the line-items header
        if any(c in ("Quantity", "Part Number", "Description") for c in cols):
            break
        if all(c == "" for c in cols):
            continue
        rows.append({
            "quantity":  cols[0] if len(cols) > 0 else "",
            "pipe_size": cols[1] if len(cols) > 1 else "",
            "connector": cols[2] if len(cols) > 2 else "",
            "hole_size": cols[3] if len(cols) > 3 else "",
        })
    return rows


def _opening_schedule_from_text(text):
    rows = []
    lines = text.split("\n")
    in_section = False
    for line in lines:
        stripped = line.strip()
        if "Opening Schedule" in stripped:
            in_section = True
            continue
        if not in_section:
            continue
        if re.match(r"^\s*Quantity\s+Part Number", stripped):
            break
        # Match:  1   24" Vitrified Clay   In 84 IN CIP Base   N/A
        m = re.match(r"^\s*(\d+)\s+(.+?)\s{2,}(.+?)\s{2,}(\S.*?)\s*$", stripped)
        if m:
            rows.append({
                "quantity":  m.group(1),
                "pipe_size": m.group(2).strip(),
                "connector": m.group(3).strip(),
                "hole_size": m.group(4).strip(),
            })
    return rows


# ---------------------------------------------------------------------------
# Line items
# ---------------------------------------------------------------------------

_KNOWN_CATEGORIES = re.compile(
    r"^(Frame\s*&\s*Ring|Joint\s*Seal|Miscellaneous|Precast|Hardware|Concrete|Accessories)\b",
    re.IGNORECASE,
)

# Matches any word-group category prefix at the start of a line, e.g. "Connectors 1 ..."
_ANY_CATEGORY_HEADER = re.compile(
    r"^([A-Za-z][A-Za-z&]*(?:\s+[A-Za-z&]+)*)\s+(?=\d)",
    re.IGNORECASE,
)


_CATEGORY_ALIASES = {"depth plus": "Precast"}


def _dedouble(s):
    if len(s) >= 4 and len(s) % 2 == 0 and all(s[i] == s[i + 1] for i in range(0, len(s), 2)):
        return s[::2]
    return s


def _normalize_category(cat):
    deduped = _dedouble(cat)
    return _CATEGORY_ALIASES.get(deduped.lower(), deduped)


def _parse_line_items(tables, text):
    # Try table extraction first
    for table in (tables or []):
        items = _line_items_from_table(table)
        if items:
            return items

    # Fallback: text-based state machine
    return _line_items_from_text(text)


def _line_items_from_table(table):
    if not table:
        return []
    flat = [str(c).strip() for row in table for c in (row or []) if c]
    if not any(_KNOWN_CATEGORIES.match(s) for s in flat):
        return []
    # Packed format: header labels crammed into one cell — fall back to text
    if any("Quantity" in s and "Part Number" in s for s in flat):
        return []

    items = []
    current_category = ""
    for row in table:
        cols = [str(c).strip() if c else "" for c in row]

        # Skip header and total rows
        if any(c in ("Quantity", "Part Number", "Description", "Weight (lbs)") for c in cols):
            continue
        if any("Total Precast Weight" in c for c in cols):
            continue
        if all(c == "" for c in cols):
            continue

        # If first column is a category name the row starts a new group
        if cols[0] and _KNOWN_CATEGORIES.match(cols[0]):
            current_category = _normalize_category(cols[0].strip())
            qty, part, desc, weight = (
                cols[1] if len(cols) > 1 else "",
                cols[2] if len(cols) > 2 else "",
                cols[3] if len(cols) > 3 else "",
                cols[4] if len(cols) > 4 else "",
            )
        else:
            qty, part, desc, weight = (
                cols[0] if len(cols) > 0 else "",
                cols[1] if len(cols) > 1 else "",
                cols[2] if len(cols) > 2 else "",
                cols[3] if len(cols) > 3 else "",
            )

        if not qty and not desc:
            continue

        items.append({
            "category":    current_category,
            "quantity":    qty,
            "part_number": part,
            "description": desc,
            "weight_lbs":  weight.replace(",", "") if weight else "",
        })
    return items


def _line_items_from_text(text, start_in_items=False):
    items = []
    current_category = ""
    in_items = start_in_items

    for raw_line in text.split("\n"):
        line = raw_line.strip()

        if not in_items:
            if _KNOWN_CATEGORIES.match(line) or _ANY_CATEGORY_HEADER.match(line):
                in_items = True
            else:
                continue

        if "Total Precast Weight" in line or "Opening Schedule" in line:
            break

        # Strip leading category name if present (known or any word-group before a quantity)
        cat_m = _KNOWN_CATEGORIES.match(line) or _ANY_CATEGORY_HEADER.match(line)
        if cat_m:
            current_category = _normalize_category(cat_m.group(1).strip())
            line = line[cat_m.end():].strip()

        # First token must be a quantity
        m = re.match(r"^(\d+)\s+(.+)$", line)
        if not m:
            continue
        qty = m.group(1)
        rest = m.group(2).strip()

        # Strip trailing weight (digits with optional comma) for Precast rows
        weight = ""
        if current_category.lower() == "precast":
            w_m = re.search(r"\s+([\d,]+)\s*$", rest)
            if w_m:
                weight = w_m.group(1).replace(",", "")
                rest = rest[:w_m.start()].strip()

        # First whitespace-free token that contains a letter = part number
        # Guard against measurement tokens like 30", 3/4"
        part = ""
        p_m = re.match(r'^([A-Za-z0-9][A-Za-z0-9.\-/]+)\s+(.+)$', rest)
        if p_m and not re.match(r'^\d+["\'/]', p_m.group(1)):
            part = p_m.group(1)
            desc = p_m.group(2).strip()
        else:
            desc = rest

        # Fallback: recover weight garbled into a description word when a long
        # description overflows into the weight column and pdfplumber interleaves
        # the characters (e.g. "TROUGHING" + "6,100" → "TROUG6H,I1N0G0")
        if not weight and current_category.lower() == "precast":
            for token in re.findall(r'[\w,]+', desc):
                if re.search(r'[A-Za-z]', token) and re.search(r'\d', token):
                    digits_only = re.sub(r'[A-Za-z]', '', token)
                    if re.match(r'^\d{1,6}(,\d{3})+$', digits_only):
                        weight = digits_only.replace(',', '')
                        break

        items.append({
            "category":    current_category,
            "quantity":    qty,
            "part_number": part,
            "description": desc,
            "weight_lbs":  weight,
        })

    return items


# ---------------------------------------------------------------------------
# Total weight
# ---------------------------------------------------------------------------

def _parse_total_weight(text):
    m = re.search(r"Total Precast Weight\s+([\d,]+)", text)
    return int(m.group(1).replace(",", "")) if m else None


# ---------------------------------------------------------------------------
# Subprocess-safe wrappers (isolates pdfplumber native crashes)
# ---------------------------------------------------------------------------

def _run_pdf_in_subprocess(func_name, pdf_bytes, timeout=60):
    import subprocess
    import pickle
    import sys
    import os
    module_dir = os.path.dirname(os.path.abspath(__file__))
    code = (
        f"import sys, pickle; sys.path.insert(0, {repr(module_dir)}); "
        f"from parse_bom_pdf import {func_name}; "
        f"data = sys.stdin.buffer.read(); "
        f"sys.stdout.buffer.write(pickle.dumps({func_name}(data)))"
    )
    proc = subprocess.run(
        [sys.executable, "-c", code],
        input=pdf_bytes,
        capture_output=True,
        timeout=timeout,
    )
    if proc.returncode != 0:
        stderr = proc.stderr.decode(errors="replace")[:300]
        raise RuntimeError(f"PDF parser crashed (exit {proc.returncode}): {stderr}")
    return pickle.loads(proc.stdout)


def parse_bom_pdf_safe(pdf_bytes):
    return _run_pdf_in_subprocess("parse_bom_pdf", pdf_bytes)


def parse_bom_by_structure_pdf_safe(pdf_bytes):
    return _run_pdf_in_subprocess("parse_bom_by_structure_pdf", pdf_bytes)


def parse_shop_drawing_pdf(pdf_bytes):
    """Parse a Shop Drawing PDF for structure names and any embedded parts schedule.

    Returns {header, structures, skipped_structures} with the same shape as
    parse_bom_by_structure_pdf.  If no parts table is found for a structure,
    one blank line item is emitted so the structure appears in the output.
    """
    try:
        import pdfplumber
    except ImportError:
        raise ImportError("pdfplumber is required: pip install pdfplumber")

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        header = _empty_header()
        if pdf.pages:
            first_text = pdf.pages[0].extract_text(layout=True) or ""
            header["release_date"] = _parse_release_date(pdf.pages[0])
            for line in first_text.split("\n")[:15]:
                stripped = line.strip()
                if stripped and len(stripped) > 4 and not re.match(r"^\d", stripped):
                    header["job_name"] = stripped
                    break

        structures = {}
        structure_order = []

        for page in pdf.pages:
            text = page.extract_text(layout=True) or ""
            structure_name = None
            for line in text.split("\n"):
                stripped = line.strip()
                if not stripped:
                    continue
                if not any(p.search(stripped) for p in _NON_STRUCTURE_PATTERNS) and \
                        any(p.search(stripped) for p in _STRUCTURE_PATTERNS):
                    structure_name = re.split(r"\s+Ship Date", stripped, flags=re.IGNORECASE)[0].strip()
                    break

            if not structure_name:
                continue

            if structure_name not in structures:
                structures[structure_name] = []
                structure_order.append(structure_name)

            for table in (page.extract_tables() or []):
                items = _line_items_from_table(table)
                if items:
                    structures[structure_name].extend(items)
                    break

        result_structures = []
        for name in structure_order:
            items = structures[name] or [{
                "category": "", "quantity": "", "part_number": "",
                "description": "", "weight_lbs": "",
            }]
            result_structures.append({
                "structure_name":       name,
                "line_items":           items,
                "total_precast_weight": None,
            })

    return {"header": header, "structures": result_structures, "skipped_structures": []}


def parse_shop_drawing_pdf_safe(pdf_bytes):
    return _run_pdf_in_subprocess("parse_shop_drawing_pdf", pdf_bytes)


def parse_ncf_docx(docx_bytes):
    """Parse a New Customer Form or Existing Customer Form .docx file.

    Returns {location, latitude, longitude, job_name, customer} where location
    is 'City, State', latitude/longitude are decimal strings (empty if not in form).
    """
    import zipfile
    import xml.etree.ElementTree as ET

    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as z:
        with z.open("word/document.xml") as f:
            root = ET.parse(f).getroot()

    parts = []
    for para in root.iter(f"{{{W}}}p"):
        texts = [t.text for t in para.iter(f"{{{W}}}t") if t.text]
        if texts:
            parts.append("".join(texts))

    # Join with double-space so label:  value patterns survive across table cells
    text = "  ".join(parts)

    def extract(label):
        m = re.search(
            rf"{re.escape(label)}\s*:?\s+(.+?)(?:\s{{2,}}|$)",
            text,
            re.IGNORECASE,
        )
        return m.group(1).strip() if m else ""

    def abbr(s):
        return _STATE_ABBR.get(s.lower(), s)

    def clean_city(s):
        return re.sub(r"\s*,?\s*\w+ County.*$", "", s, flags=re.IGNORECASE).strip()

    job_site_city  = extract("Job Site City")
    job_site_state = extract("Job Site State")

    if job_site_city and job_site_state:
        city    = clean_city(job_site_city)
        state   = abbr(job_site_state)
        zipcode = extract("Job Site Zip Code")
        location_source = "NCF Job Site"
    else:
        billing_city  = extract("Billing City")
        billing_state = extract("Billing State")
        city    = clean_city(job_site_city or billing_city)
        state   = abbr(job_site_state or billing_state)
        zipcode = ""
        location_source = "NCF Billing" if (billing_city or billing_state) else ""

    location = f"{city}, {state}" if city and state else (city or state)

    return {
        "location":         location,
        "location_source":  location_source,
        "zipcode":          zipcode,
        "job_name":         extract("Name of Job"),
        "customer":         extract("Business Name of Armorock Customer"),
    }
