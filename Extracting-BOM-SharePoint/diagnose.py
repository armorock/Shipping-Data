"""
Diagnostic script — parses all PDFs in Data/ and writes a report to output/diagnosis.txt.
Run this before sharing output with Claude so it has a concrete view of what's being extracted.

Usage:
    python diagnose.py
"""

import os, sys, re, io
sys.path.insert(0, os.path.dirname(__file__))

from parse_bom_pdf import parse_bom_by_structure_pdf, parse_bom_pdf

DATA_DIR   = os.path.join(os.path.dirname(__file__), "Data")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "output")
REPORT     = os.path.join(OUTPUT_DIR, "diagnosis.txt")

os.makedirs(OUTPUT_DIR, exist_ok=True)

lines = []
def log(s=""):
    lines.append(s)
    print(s)


def _is_by_structure(filename):
    return "bom by structure" in filename.lower() and filename.lower().endswith(".pdf")

def _is_summary(filename):
    return "bom summary" in filename.lower() and filename.lower().endswith(".pdf")


for filename in sorted(os.listdir(DATA_DIR)):
    if not filename.lower().endswith(".pdf"):
        continue

    path = os.path.join(DATA_DIR, filename)
    with open(path, "rb") as f:
        pdf_bytes = f.read()

    log(f"{'='*70}")
    log(f"FILE: {filename}")
    log(f"SIZE: {len(pdf_bytes):,} bytes")

    try:
        if _is_by_structure(filename):
            result = parse_bom_by_structure_pdf(pdf_bytes)
            log(f"TYPE: BOM by Structure")
            log(f"JOB:  {result['header'].get('job_name','')}  |  Released: {result['header'].get('release_date','')}")
            log(f"STRUCTURES: {len(result['structures'])}")
            for s in result["structures"]:
                items     = s["line_items"]
                precast   = [i for i in items if i["category"].lower() == "precast"]
                no_weight = [i for i in precast if not i["weight_lbs"]]
                log(f"  [{s['structure_name']}]  {len(items)} items  ({len(precast)} precast)")
                for i in items:
                    w = f"  weight={i['weight_lbs']}" if i["category"].lower() == "precast" else ""
                    flag = "  *** NO WEIGHT ***" if i["category"].lower() == "precast" and not i["weight_lbs"] else ""
                    log(f"    [{i['category']:15s}] {i['part_number']:25s} qty={i['quantity']}{w}{flag}")
                if no_weight:
                    log(f"  WARNING: {len(no_weight)} precast item(s) missing weight")

        elif _is_summary(filename):
            result = parse_bom_pdf(pdf_bytes)
            log(f"TYPE: BOM Summary")
            log(f"JOB:  {result['header'].get('job_name','')}  |  Released: {result['header'].get('release_date','')}")
            items   = result["line_items"]
            precast = [i for i in items if i["category"].lower() == "precast"]
            log(f"ITEMS: {len(items)}  ({len(precast)} precast)")
            for i in items:
                w = f"  weight={i['weight_lbs']}" if i["category"].lower() == "precast" else ""
                flag = "  *** NO WEIGHT ***" if i["category"].lower() == "precast" and not i["weight_lbs"] else ""
                log(f"  [{i['category']:15s}] {i['part_number']:25s} qty={i['quantity']}{w}{flag}")

        else:
            log("TYPE: unknown — skipped")

    except Exception as exc:
        log(f"ERROR: {exc}")

    log()

with open(REPORT, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))

print(f"\nReport written to: {REPORT}")
