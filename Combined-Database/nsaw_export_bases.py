"""
NSAW data prep — exports every shipped base with its Gen4 canonical item name and job code.
Output: Combined-Database/output/bases_by_job_code.csv

Re-run this script before each NSAW to refresh the dataset from the latest pieces_ledger.
"""
import os
import re
import pandas as pd

LEDGER = "Combined-Database/output/pieces_ledger.csv"
OUTPUT = "Combined-Database/output/bases_by_job_code.csv"


# ---------------------------------------------------------------------------
# Part number → Gen4 name converter (handles Gen1 / Gen2 / Gen3 / Gen4)
# Logic ported from /item-name-detail-interpreter + /part-name-builder skills
# ---------------------------------------------------------------------------

def _norm_tr(t):
    if not t:
        return ""
    t = str(t).strip()
    if re.match(r'^1\.3{1,2}$', t):
        return ".133"
    if re.match(r'^1(\.0?)?$', t):
        return ".1"
    if t.startswith("0."):
        return t[1:]
    return t

def _build_mhb(dia, ht, tr="", es=False, de=0, wall=""):
    if not dia or not ht:
        return ""
    tr_sfx = "FF" if tr == "FF" else _norm_tr(tr)
    es_sfx = "ES" if es else ""
    de_sfx = "/DE2" if de == 2 else "/DE" if de == 1 else ""
    w_sfx  = f"-{wall}" if wall else ""
    return f"MHB{dia}{ht}{tr_sfx}{es_sfx}{de_sfx}{w_sfx}"

def _parse_gen1(name):
    # No-type-letter base: {dia}{1-digit-ft}{decimal} e.g. 6041.33
    m = re.match(r'^(144|120|96|84|72|60|48)(\d)(1?\.\d+)$', name)
    if m:
        return _build_mhb(m[1], str(int(m[2]) * 12), tr=m[3])

    # Dash format: {dia}{ft}-{in}BP?{.tr?}
    m = re.match(r'^(144|120|96|84|72|60|48)(\d)-(\d+)BP?(\d*\.\d+)?$', name, re.I)
    if m:
        ht = str(int(m[2]) * 12 + int(m[3]))
        return _build_mhb(m[1], ht, tr=m[4] or "")

    # BFF flat-floor
    m = re.match(r'^(144|120|96|84|72|60|48)(\d+)BFF$', name, re.I)
    if m:
        ht = str(int(m[2]) * 12) if len(m[2]) == 1 else m[2]
        return _build_mhb(m[1], ht, tr="FF")

    # B / BP with optional troughing decimal
    m = re.match(r'^(144|120|96|84|72|60|48)(\d+)BP?(\d*\.\d+)?$', name, re.I)
    if m:
        ht = str(int(m[2]) * 12) if len(m[2]) == 1 else m[2]
        return _build_mhb(m[1], ht, tr=m[3] or "")

    return ""

def _parse_gen2(name):
    if re.match(r'^60\d{3}-', name):
        name = "7" + name[1:]
    m = re.match(r'^7(\d{2})(\d{2})-(.+)$', name)
    if not m:
        return ""
    dia  = str(int(m[1]) * 12)
    ht   = m[2]
    sfx  = m[3].upper()
    es   = sfx.endswith("ES")
    base = sfx[:-2] if es else sfx
    tr_map = {
        "B": "", "B1": ".1", "B50": ".5", "B75": ".75",
        "B100": ".1", "B116": "1.16", "B133": ".133", "BFF": "FF",
    }
    tr = tr_map.get(base, "")
    return _build_mhb(dia, ht, tr=tr, es=es)

def _parse_gen3(name):
    prefix = name[0].upper()
    m = re.match(r'^[BFTMV](\d{2})(\d{2})', name, re.I)
    if not m:
        return ""
    dia = str(int(m[1]) * 12)
    ht  = m[2]
    tr  = "FF" if prefix == "F" else ""
    return _build_mhb(dia, ht, tr=tr)

def _parse_gen4(name):
    m = re.match(
        r'^MHB(\d{2,3})(\d{2,3})(FF|\.\d+)?(ES)?(/DE2|/DE)?(-\d)?$',
        name, re.I
    )
    if not m:
        return name  # Return unchanged; unusual suffix but still Gen4
    dia, ht = m[1], m[2]
    tr       = (m[3] or "").upper()
    es       = bool(m[4])
    de       = 2 if m[5] == "/DE2" else (1 if m[5] == "/DE" else 0)
    wall     = m[6][1:] if m[6] else ""
    return _build_mhb(dia, ht, tr=tr, es=es, de=de, wall=wall)

def to_gen4_name(raw):
    if not isinstance(raw, str) or not raw.strip():
        return ""
    name = raw.strip().replace(" (inactive)", "").replace("(inactive)", "").strip()

    chk = "7" + name[1:] if re.match(r'^60\d{3}-', name) else name

    if re.match(r'^MHB', chk, re.I):
        return _parse_gen4(chk)
    if re.match(r'^(MH|RMH|BOX|MT)', chk, re.I):
        return name  # Non-base Gen4 type — return as-is (shouldn't appear for BASE rows)
    if re.match(r'^[57]\d{4}-', chk):
        return _parse_gen2(chk)
    if re.match(r'^[BFTMV]\d{4}', chk, re.I):
        return _parse_gen3(chk)
    if re.match(r'^(144|120|96|84|72|60|48)', chk):
        return _parse_gen1(chk)

    return ""  # Unrecognized / custom / material


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

ledger = pd.read_csv(LEDGER)
bases  = ledger[ledger["part_type"] == "BASE"].copy()

# Keep standard 3-letter A–E job codes, plus blank/null (pre-2018 jobs with no job code assigned)
valid_jc = bases["job_code"].isna() | bases["job_code"].astype(str).str.match(r'^[A-Ea-e][A-Za-z]{2}$')
bases = bases[valid_jc].copy()

bases["item_name"] = bases["part_number"].apply(to_gen4_name)

cols = ["part_number", "item_name", "job_code", "ship_date", "year", "city", "state", "zip", "plant"]
bases = bases[cols].sort_values(["job_code", "ship_date"]).reset_index(drop=True)

os.makedirs("Combined-Database/output", exist_ok=True)
bases.to_csv(OUTPUT, index=False)

matched = (bases["item_name"] != "").sum()
print(f"Wrote {len(bases):,} rows to {OUTPUT}")
print(f"item_name populated: {matched:,} / {len(bases):,} ({matched/len(bases)*100:.1f}%)")
