"""
Read output/skipped_structures.csv and group by leading pattern so we can
identify what structure name formats need to be added to _STRUCTURE_PATTERNS.

Run: python analyze_skips.py
"""
import csv
import re
from collections import Counter, defaultdict

CSV_PATH = "output/skipped_structures.csv"

rows = []
with open(CSV_PATH, newline="", encoding="utf-8") as f:
    rows = list(csv.DictReader(f))

print(f"Total skipped: {len(rows)}\n")

prefix_groups = defaultdict(list)
for row in rows:
    name = row["structure_name"]
    m = re.match(r"^([A-Za-z]+)", name)
    prefix = m.group(1).upper() if m else "(non-alpha)"
    prefix_groups[prefix].append(name)

print(f"{'Prefix':<15} {'Count':>6}  {'Examples'}")
print("-" * 70)
for prefix, names in sorted(prefix_groups.items(), key=lambda x: -len(x[1])):
    examples = ", ".join(f'"{n}"' for n in sorted(set(names))[:3])
    print(f"{prefix:<15} {len(names):>6}  {examples}")
