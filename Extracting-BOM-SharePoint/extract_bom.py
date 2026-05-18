"""
Extract BOM data from Armorock SharePoint job folders.

Folder naming convention: {3-LETTER-CODE}-{optional space}{Job Name}
  e.g.  CVG- Farmersville Plants #3
        CVP-North River Ranch IV-c1

Outputs:
  bom_line_items.csv  — one row per BOM line item, keyed by job_code
  bom_openings.csv    — one row per pipe opening, keyed by job_code
"""

import csv
import re
import sys
from graph_client import acquire_token, graph_get_all, GRAPH_ROOT
from sharepoint_client import get_site, get_drive, list_children, iter_files, download_file
from parse_bom_pdf import parse_bom_pdf

# ---------------------------------------------------------------------------
# Configure for your SharePoint environment
# ---------------------------------------------------------------------------
HOSTNAME   = "armorockllc.sharepoint.com"
SITE_PATH  = "/sites/JobData2026"  # server-relative path to the site
DRIVE_NAME = "Job Data 2026"       # document library name
JOBS_ROOT  = "root"                # job folders sit at the drive root
OUTPUT_LINE_ITEMS = "bom_line_items.csv"
OUTPUT_OPENINGS   = "bom_openings.csv"
# ---------------------------------------------------------------------------


_JOB_FOLDER_RE = re.compile(r"^([A-Z]{3})\s*-\s*(.+)$")


def parse_folder_name(name):
    """Return (job_code, job_name) from a folder like 'CVG- Farmersville Plants #3'."""
    m = _JOB_FOLDER_RE.match(name.strip())
    if m:
        return m.group(1), m.group(2).strip()
    return None, name.strip()


def is_bom_pdf(filename):
    name = filename.lower()
    return "bom" in name and name.endswith(".pdf")


def flatten_line_items(job_code, job_name, bom):
    header = bom["header"]
    base = {
        "job_code":            job_code,
        "job_name_folder":     job_name,
        "pdf_job_name":        header["job_name"],
        "pdf_job_number":      header["job_number"],
        "location":            header["location"],
        "contractor":          header["contractor"],
        "total_precast_weight": bom["total_precast_weight"] or "",
    }
    for item in bom["line_items"]:
        yield {**base, **item}


def flatten_openings(job_code, job_name, bom):
    header = bom["header"]
    base = {
        "job_code":        job_code,
        "job_name_folder": job_name,
        "location":        header["location"],
    }
    for op in bom["opening_schedule"]:
        yield {**base, **op}


def write_csv(path, rows):
    if not rows:
        print(f"  (no rows to write for {path})")
        return
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)
    print(f"Wrote {len(rows):,} rows -> {path}")


def main():
    token = acquire_token()

    site = get_site(token, HOSTNAME, SITE_PATH)
    site_id = site["id"]
    print(f"Site: {site['displayName']}")

    drive = get_drive(token, site_id, DRIVE_NAME)
    drive_id = drive["id"]
    print(f"Drive: {drive['name']}")

    # Get top-level job folders
    children = list_children(token, drive_id, JOBS_ROOT)
    job_folders = [c for c in children if "folder" in c]
    print(f"Found {len(job_folders)} top-level folders\n")

    line_item_rows = []
    opening_rows = []
    skipped = []
    errors = []

    for folder in job_folders:
        job_code, job_name = parse_folder_name(folder["name"])
        if not job_code:
            skipped.append(folder["name"])
            continue

        bom_files = [
            item for item in iter_files(token, drive_id, folder["id"], recursive=True)
            if is_bom_pdf(item["name"])
        ]

        if not bom_files:
            print(f"[{job_code}] No BOM PDF found — skipping")
            continue

        for file_item in bom_files:
            print(f"[{job_code}] Parsing: {file_item['name']}")
            try:
                pdf_bytes = download_file(token, drive_id, file_item["id"])
                bom = parse_bom_pdf(pdf_bytes)
                line_item_rows.extend(flatten_line_items(job_code, job_name, bom))
                opening_rows.extend(flatten_openings(job_code, job_name, bom))
            except Exception as exc:
                msg = f"[{job_code}] ERROR {file_item['name']}: {exc}"
                print(f"  {msg}", file=sys.stderr)
                errors.append(msg)

    print()
    write_csv(OUTPUT_LINE_ITEMS, line_item_rows)
    write_csv(OUTPUT_OPENINGS, opening_rows)

    if skipped:
        print(f"\nSkipped {len(skipped)} folders with no job code: {skipped[:5]}{'...' if len(skipped) > 5 else ''}")
    if errors:
        print(f"\n{len(errors)} errors encountered — see stderr output above")


if __name__ == "__main__":
    main()
