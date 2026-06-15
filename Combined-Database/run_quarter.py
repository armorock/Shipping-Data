import argparse
from datetime import datetime

import sources
import trust_model
import build_observation_db
import build_registry
import build_pieces_ledger
import build_reports
import snapshot_and_diff


def default_quarter():
    now = datetime.now()
    return f"{now.year}-Q{(now.month - 1) // 3 + 1}"


def main():
    ap = argparse.ArgumentParser(description="Run the Combined-Database pipeline for a quarter.")
    ap.add_argument("--quarter", default=default_quarter(),
                    help="Quarter label, e.g. 2026-Q2 (defaults to current quarter).")
    ap.add_argument("--no-snapshot", action="store_true", help="Skip snapshot/diff at the end.")
    args = ap.parse_args()

    sources.ensure_dirs()
    print(f"=== Combined-Database run: {args.quarter} ===")
    print("[1/6] trust model");        trust_model.write_report()
    print("[2/6] observation DB");     build_observation_db.build()
    print("[3/6] project registry");   build_registry.build()
    print("[4/6] pieces ledger");      build_pieces_ledger.build()
    print("[5/6] reports & review");   build_reports.build()
    if args.no_snapshot:
        print("[6/6] snapshot skipped")
    else:
        print("[6/6] snapshot & diff"); snapshot_and_diff.snapshot(args.quarter)
    print("=== done ===")


if __name__ == "__main__":
    main()
