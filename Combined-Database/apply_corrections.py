import csv
import glob
import os
import sqlite3
import sys
from datetime import datetime

import openpyxl

import sources
import resolutions
import build_registry
import build_pieces_ledger
import build_reports


def _winner(con, ek, field):
    r = con.execute("SELECT winner FROM entity_field_resolution WHERE entity_key=? AND field=?",
                    (ek, field)).fetchone()
    return r[0] if r else None


def read_confirmations(files):
    """Yield (entity_key, field, chosen_value) from the Confirmed Value column of each review xlsx."""
    for path in files:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        rows = ws.iter_rows(values_only=True)
        hdr = [str(c).strip() if c is not None else "" for c in next(rows)]
        idx = {name: i for i, name in enumerate(hdr)}
        jc_i, f_i, cv_i = idx.get("job_code"), idx.get("field"), idx.get("Confirmed Value")
        if cv_i is None or jc_i is None or f_i is None:
            wb.close()
            continue
        for row in rows:
            val = row[cv_i]
            if val is None or str(val).strip() == "":
                continue
            yield str(row[jc_i]).strip(), str(row[f_i]).strip(), str(val).strip(), os.path.basename(path)
        wb.close()


def apply(files):
    sources.ensure_dirs()
    if not os.path.exists(sources.OBSERVATIONS_DB):
        print("observations.db missing; run build_observation_db.py first")
        return
    con = sqlite3.connect(sources.OBSERVATIONS_DB)
    log_rows = []
    upserts = []
    for ek, field, val, src in read_confirmations(files):
        old = _winner(con, ek, field)
        upserts.append((ek, field, val, f"review:{src}"))
        log_rows.append([ek, field, old, val, src])
    con.close()

    if not upserts:
        print("no confirmations found in:", [os.path.basename(f) for f in files])
        return
    n = resolutions.upsert(upserts)
    print(f"recorded {n} confirmed resolutions")

    log_path = os.path.join(sources.OUTPUT_DIR, "corrections_log.csv")
    exists = os.path.exists(log_path)
    with open(log_path, "a", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        if not exists:
            w.writerow(["applied_at", "entity_key", "field", "old_value", "new_value", "source_list"])
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for r in log_rows:
            w.writerow([stamp] + r)

    print("rebuilding registry, ledger, reports...")
    build_registry.build()
    build_pieces_ledger.build()
    build_reports.build()
    print("done. corrections_log:", log_path)


def main():
    args = sys.argv[1:]
    if args:
        files = [a if os.path.isabs(a) else os.path.join(sources.REVIEW_DIR, a) for a in args]
    else:
        files = sorted(glob.glob(os.path.join(sources.REVIEW_DIR, "*.xlsx")))
    files = [f for f in files if os.path.exists(f)]
    if not files:
        print("no review files found")
        return
    apply(files)


if __name__ == "__main__":
    main()
