"""
Build a single unioned XLSM workbook from both BOM extraction projects.

Source split (no overlap):
  M Drive          2016-2022  extracting M drive/output/
  BOM Str. Detail  2023-2026  BOM Structure Detail/output/

Output: output\all_bom_union.xlsm

All values written as explicit strings so Excel cannot auto-convert
date-like product numbers (e.g. "12-08") to dates.
"""
import csv
import os
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

MDRIVE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
BSD_DIR    = r"C:\Users\JohnLeitzke\Code\BOM Structure Detail\output"
OUTPUT_DIR = MDRIVE_DIR
OUT_FILE   = os.path.join(OUTPUT_DIR, "all_bom_union.xlsm")

SOURCES = [
    (MDRIVE_DIR, range(2016, 2023)),   # M Drive: 2016-2022
    (BSD_DIR,    range(2023, 2027)),   # BOM Structure Detail: 2023-2026
]


def _iter_rows(directory, years):
    for year in years:
        path = os.path.join(directory, f"all_bom_{year}.csv")
        if not os.path.exists(path):
            print(f"  {year}: not found in {directory} — skipping")
            continue
        with open(path, encoding="utf-8", newline="") as f:
            reader = csv.reader(f)
            next(reader)  # skip header
            count = 0
            for row in reader:
                yield row
                count += 1
        print(f"  {year}: {count:,} rows")


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    wb = Workbook(write_only=True)
    ws = wb.create_sheet("BOM 2016-2026")

    # Write header from first available file
    header_written = False
    for directory, years in SOURCES:
        if header_written:
            break
        for year in years:
            path = os.path.join(directory, f"all_bom_{year}.csv")
            if os.path.exists(path):
                with open(path, encoding="utf-8", newline="") as f:
                    header = next(csv.reader(f))
                ws.append(header)
                header_written = True
                break

    total = 0
    for label, (directory, years) in zip(("M Drive 2016-2022", "BOM Str. Detail 2023-2026"), SOURCES):
        print(f"\n{label}:")
        for row in _iter_rows(directory, years):
            ws.append(row)
            total += 1

    wb.save(OUT_FILE)
    print(f"\nWrote {total:,} rows -> {OUT_FILE}")


if __name__ == "__main__":
    main()
