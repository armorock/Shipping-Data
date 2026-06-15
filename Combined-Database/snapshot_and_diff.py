import csv
import os
import shutil
from collections import Counter

import openpyxl

import sources

ARTIFACTS = ["pieces_ledger.csv", "project_registry.csv", "rehab_nonbase_report.xlsx",
             "quoted_not_shipped.xlsx", "structure_completeness_roadmap.xlsx",
             "project_provenance.xlsx", "dispatch_vs_erp_reconciliation.xlsx", "trust_report.md"]
CLASSES = ["base", "rehab", "non-base", "unknown"]


def _metrics():
    led = list(csv.DictReader(open(os.path.join(sources.OUTPUT_DIR, "pieces_ledger.csv"), encoding="utf-8")))
    reg = list(csv.DictReader(open(os.path.join(sources.OUTPUT_DIR, "project_registry.csv"), encoding="utf-8")))
    cls = Counter(p["structure_class"] for p in led)
    m = {"pieces": len(led), "job_codes": len(reg),
         "needs_review": sum(1 for r in reg if r.get("needs_review") == "Y")}
    for c in CLASSES:
        m[c] = cls[c]
    return m, {p["piece_id"] for p in led}, led


def _prev_quarter(quarter):
    if not os.path.exists(sources.SNAPSHOT_DIR):
        return None
    others = sorted(d for d in os.listdir(sources.SNAPSHOT_DIR)
                    if d != quarter and os.path.isdir(os.path.join(sources.SNAPSHOT_DIR, d)))
    return others[-1] if others else None


def snapshot(quarter):
    sources.ensure_dirs()
    m, piece_ids, led = _metrics()

    dest = os.path.join(sources.SNAPSHOT_DIR, quarter)
    os.makedirs(dest, exist_ok=True)
    for a in ARTIFACTS:
        src = os.path.join(sources.OUTPUT_DIR, a)
        if os.path.exists(src):
            shutil.copy(src, os.path.join(dest, a))

    metrics_path = os.path.join(sources.OUTPUT_DIR, "snapshot_metrics.csv")
    cols = ["quarter", "pieces"] + CLASSES + ["job_codes", "needs_review"]
    existing = {}
    if os.path.exists(metrics_path):
        for r in csv.DictReader(open(metrics_path, encoding="utf-8")):
            existing[r["quarter"]] = r
    existing[quarter] = {"quarter": quarter, **{k: m[k] for k in cols if k != "quarter"}}
    with open(metrics_path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        for q in sorted(existing):
            w.writerow(existing[q])

    prev = _prev_quarter(quarter)
    if not prev:
        print(f"snapshot {quarter}: {m} (no prior snapshot to diff)")
        return
    prev_led = list(csv.DictReader(open(os.path.join(sources.SNAPSHOT_DIR, prev, "pieces_ledger.csv"),
                                        encoding="utf-8")))
    prev_ids = {p["piece_id"] for p in prev_led}
    new_ids = piece_ids - prev_ids
    gone_ids = prev_ids - piece_ids
    prev_cls = Counter(p["structure_class"] for p in prev_led)

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "summary"
    ws.append(["metric", prev, quarter, "delta"])
    pm = {"pieces": len(prev_led), **{c: prev_cls[c] for c in CLASSES}}
    for k in ["pieces"] + CLASSES:
        ws.append([k, pm.get(k, 0), m[k], m[k] - pm.get(k, 0)])
    ws.append(["new_pieces", "", len(new_ids), ""])
    ws.append(["removed_pieces", "", len(gone_ids), ""])
    ws2 = wb.create_sheet("new_pieces")
    ws2.append(["piece_id", "job_code", "year", "part_number", "structure_class", "state"])
    by_id = {p["piece_id"]: p for p in led}
    for pid in sorted(new_ids):
        p = by_id[pid]
        ws2.append([pid, p["job_code"], p["year"], p["part_number"], p["structure_class"], p["state"]])
    for w in (ws, ws2):
        w.freeze_panes = "A2"
    path = os.path.join(sources.OUTPUT_DIR, f"changes_{quarter}.xlsx")
    wb.save(path)
    print(f"snapshot {quarter}: {m}")
    print(f"diff vs {prev}: +{len(new_ids)} pieces, -{len(gone_ids)} -> {os.path.basename(path)}")


if __name__ == "__main__":
    import sys
    q = sys.argv[1] if len(sys.argv) > 1 else "manual"
    snapshot(q)
