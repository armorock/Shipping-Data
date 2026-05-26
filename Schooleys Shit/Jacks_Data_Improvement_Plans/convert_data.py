"""
Converts raw exports in data/ into clean CSVs ready for jl_build_jobcode_db.py.

  Notion ZIP  →  data/notion_export.csv
  NetSuite XLS (SpreadsheetML XML)  →  data/netsuite_jobs.csv

Run: python convert_data.py
"""

import os, re, csv, io, glob, zipfile
import xml.etree.ElementTree as ET

BASE       = os.path.dirname(os.path.abspath(__file__))
DATA_DIR   = os.path.join(BASE, "data")

NOTION_OUT   = os.path.join(DATA_DIR, "notion_export.csv")
NETSUITE_OUT = os.path.join(DATA_DIR, "netsuite_jobs.csv")

NOTION_KEEP  = ["Job Code", "City", "State", "County", "Bidding Contractors", "Job Status"]
NS_XML       = "urn:schemas-microsoft-com:office:spreadsheet"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def find_file(pattern):
    matches = glob.glob(os.path.join(DATA_DIR, pattern))
    return matches[0] if matches else None


def normalize_code(val):
    if not val:
        return None
    v = re.sub(r"\s+", "", str(val)).upper()
    m = re.match(r"^([A-E][A-Z0-9]{2})", v)
    return m.group(1) if m else None


# Items that would double-count if included in job value totals
_SKIP_AMT_ITEMS = {"subtotal", "avatax"}


# ---------------------------------------------------------------------------
# Notion: unpack nested ZIP, extract main DB CSV, keep relevant columns
# ---------------------------------------------------------------------------

def convert_notion():
    zip_path = find_file("*.zip")
    if not zip_path:
        print("  No Notion ZIP found in data/ — skipping")
        return

    print(f"Reading Notion ZIP: {os.path.basename(zip_path)}")

    outer = zipfile.ZipFile(zip_path)
    # May be double-nested
    inner_names = outer.namelist()
    if len(inner_names) == 1 and inner_names[0].endswith(".zip"):
        inner = zipfile.ZipFile(io.BytesIO(outer.read(inner_names[0])))
    else:
        inner = outer

    # Find the Armorock Main Database CSV
    target = next(
        (n for n in inner.namelist() if "armorock main database" in n.lower()),
        None
    )
    if not target:
        # Fall back to the largest CSV
        csvs = [n for n in inner.namelist() if n.endswith(".csv")]
        target = max(csvs, key=lambda n: len(inner.read(n)), default=None)

    if not target:
        print("  No CSV found inside Notion ZIP — skipping")
        return

    print(f"  Extracting: {target}")
    raw = inner.read(target).decode("utf-8-sig")
    reader = csv.DictReader(io.StringIO(raw))

    rows_written = 0
    with open(NOTION_OUT, "w", newline="", encoding="utf-8") as out:
        writer = csv.DictWriter(out, fieldnames=NOTION_KEEP)
        writer.writeheader()
        for row in reader:
            code = normalize_code(row.get("Job Code", ""))
            if not code:
                continue
            writer.writerow({
                "Job Code":            code,
                "City":                row.get("City", "").strip(),
                "State":               row.get("State", "").strip(),
                "County":              row.get("County", "").strip(),
                "Bidding Contractors": row.get("Bidding Contractors", "").strip(),
                "Job Status":          row.get("Job Status", "").strip(),
            })
            rows_written += 1

    print(f"  Wrote {rows_written:,} rows -> {NOTION_OUT}")


# ---------------------------------------------------------------------------
# NetSuite: parse SpreadsheetML XML, one row per job code
# ---------------------------------------------------------------------------

NS_OUT_FIELDS = [
    "Job Code", "Customer", "NetSuite Date",
    "Fulfilled", "Shipping City", "Shipping State", "Shipping Zip", "Job Value",
]


def convert_netsuite():
    xls_path = find_file("*.xls")
    if not xls_path:
        print("  No NetSuite XLS found in data/ — skipping")
        return

    print(f"Reading NetSuite XLS: {os.path.basename(xls_path)}")

    with open(xls_path, encoding="utf-8", errors="replace") as f:
        content = f.read()

    ns_xml = {"ss": NS_XML}
    root  = ET.fromstring(content.encode("utf-8"))
    ws    = root.find(".//ss:Worksheet", ns_xml)
    table = ws.find("ss:Table", ns_xml)
    rows  = table.findall("ss:Row", ns_xml)

    def cell_val(cell):
        d = cell.find("ss:Data", ns_xml)
        return d.text.strip() if d is not None and d.text else ""

    headers = [cell_val(c) for c in rows[0].findall("ss:Cell", ns_xml)]

    def idx(name):
        return headers.index(name) if name in headers else None

    jc_idx    = idx("Job Code")
    name_idx  = idx("Name")
    date_idx  = idx("Date")
    item_idx  = idx("Item")
    ful_idx   = idx("Fulfilled/Received (Line Level)")
    city_idx  = idx("Shipping City")
    state_idx = idx("Shipping State/Province")
    zip_idx   = idx("Shipping Zip")
    amt_idx   = idx("Amount")

    if jc_idx is None or name_idx is None:
        print("  ERROR: Job Code or Name column not found")
        return

    by_code = {}
    for row in rows[1:]:
        cells = row.findall("ss:Cell", ns_xml)
        def v(i):
            if i is None or i >= len(cells):
                return ""
            d = cells[i].find("ss:Data", ns_xml)
            return d.text.strip() if d is not None and d.text else ""

        code = normalize_code(v(jc_idx))
        if not code:
            continue

        if code not in by_code:
            by_code[code] = {
                "customer": None, "date": None,
                "fulfilled": False,
                "city": None, "state": None, "zip": None,
                "job_value": 0.0,
            }
        rec = by_code[code]

        name = v(name_idx)
        if name and not rec["customer"]:
            rec["customer"] = name

        raw_date = v(date_idx)
        if raw_date and not rec["date"]:
            rec["date"] = raw_date[:10]  # keep YYYY-MM-DD only

        if v(ful_idx).lower() == "yes":
            rec["fulfilled"] = True

        city = v(city_idx).rstrip(",").strip()
        if city and not rec["city"]:
            rec["city"] = city

        state = v(state_idx).strip()
        if state and not rec["state"]:
            rec["state"] = state[:2].upper() if len(state) > 2 else state.upper()

        zip_val = v(zip_idx).strip()
        if zip_val and not rec["zip"]:
            rec["zip"] = zip_val[:5]

        item = v(item_idx).lower()
        amt_raw = v(amt_idx)
        if amt_raw and item not in _SKIP_AMT_ITEMS:
            try:
                rec["job_value"] += float(amt_raw)
            except ValueError:
                pass

    with open(NETSUITE_OUT, "w", newline="", encoding="utf-8") as out:
        writer = csv.DictWriter(out, fieldnames=NS_OUT_FIELDS)
        writer.writeheader()
        for code in sorted(by_code):
            rec = by_code[code]
            writer.writerow({
                "Job Code":      code,
                "Customer":      rec["customer"] or "",
                "NetSuite Date": rec["date"] or "",
                "Fulfilled":     "True" if rec["fulfilled"] else "False",
                "Shipping City": rec["city"] or "",
                "Shipping State": rec["state"] or "",
                "Shipping Zip":  rec["zip"] or "",
                "Job Value":     f"{rec['job_value']:.2f}" if rec["job_value"] else "",
            })

    print(f"  Wrote {len(by_code):,} job codes -> {NETSUITE_OUT}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    os.makedirs(DATA_DIR, exist_ok=True)
    print("=== Converting Notion export ===")
    convert_notion()
    print()
    print("=== Converting NetSuite export ===")
    convert_netsuite()
    print()
    print("Done. Run jl_build_jobcode_db.py next.")
