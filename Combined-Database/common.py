import re

STATE_ABBREV = {
    "ALABAMA": "AL", "ALASKA": "AK", "ARIZONA": "AZ", "ARKANSAS": "AR",
    "CALIFORNIA": "CA", "COLORADO": "CO", "CONNECTICUT": "CT", "DELAWARE": "DE",
    "DISTRICT OF COLUMBIA": "DC", "FLORIDA": "FL", "GEORGIA": "GA", "HAWAII": "HI",
    "IDAHO": "ID", "ILLINOIS": "IL", "INDIANA": "IN", "IOWA": "IA", "KANSAS": "KS",
    "KENTUCKY": "KY", "LOUISIANA": "LA", "MAINE": "ME", "MARYLAND": "MD",
    "MASSACHUSETTS": "MA", "MICHIGAN": "MI", "MINNESOTA": "MN", "MISSISSIPPI": "MS",
    "MISSOURI": "MO", "MONTANA": "MT", "NEBRASKA": "NE", "NEVADA": "NV",
    "NEW HAMPSHIRE": "NH", "NEW JERSEY": "NJ", "NEW MEXICO": "NM", "NEW YORK": "NY",
    "NORTH CAROLINA": "NC", "NORTH DAKOTA": "ND", "OHIO": "OH", "OKLAHOMA": "OK",
    "OREGON": "OR", "PENNSYLVANIA": "PA", "RHODE ISLAND": "RI",
    "SOUTH CAROLINA": "SC", "SOUTH DAKOTA": "SD", "TENNESSEE": "TN", "TEXAS": "TX",
    "UTAH": "UT", "VERMONT": "VT", "VIRGINIA": "VA", "WASHINGTON": "WA",
    "WEST VIRGINIA": "WV", "WISCONSIN": "WI", "WYOMING": "WY",
}
STATE_CODES = set(STATE_ABBREV.values())

MISSING_TOKENS = {"", "NONE", "#MISSING", "N/A", "NA", "NULL"}


def is_blank(v):
    return v is None or (isinstance(v, str) and v.strip().upper() in MISSING_TOKENS)


def norm_city(v):
    if not isinstance(v, str) or not v.strip():
        return None
    if v.strip().startswith("="):  # VLOOKUP formula corruption in BABY
        return None
    s = " ".join(v.upper().replace(".", " ").replace(",", " ").split())
    words = s.split()
    if len(words) == 1 and words[0] in STATE_CODES:
        return None
    if len(words) >= 2 and words[-1] in STATE_CODES:
        s = " ".join(words[:-1])
    return s or None


def norm_county(v):
    if not isinstance(v, str) or not v.strip():
        return None
    s = " ".join(v.upper().replace(",", " ").split())
    words = s.split()
    if len(words) >= 2 and words[-1] in STATE_CODES:
        words = words[:-1]
    if words and words[-1] == "COUNTY":
        words = words[:-1]
    s = " ".join(words)
    return s or None


def is_county_value(v):
    if not isinstance(v, str):
        return False
    return "COUNTY" in v.upper().split()


def norm_state(v):
    if not isinstance(v, str) or not v.strip():
        return None
    s = v.strip().upper()
    if s.startswith("="):
        return None
    if s in STATE_CODES:
        return s
    return STATE_ABBREV.get(s)


def norm_zip(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        s = str(int(v)).zfill(5)
    else:
        s = str(v).strip().split("-")[0]
    return s if len(s) == 5 and s.isdigit() and s != "00000" else None


def norm_job_code(v):
    if v is None:
        return None
    s = str(v).strip()
    return s.upper() if s and s.lower() != "none" else None


def norm_part_number(v):
    if v is None:
        return None
    s = str(v).strip()
    return s if s and s.lower() not in ("none", "#missing") else None


REHAB_TYPES = {"RMH", "RMHC", "RMHL", "RMHLC"}
BASE_TYPES = {"MHB", "MHBT", "BOX", "BOXF", "BOXFF", "BOXSFF"}
NONBASE_TYPES = {
    "MHS", "MHC", "MHCC", "MHL", "MHLC", "MHTL", "MHTLC", "MHGR", "MHT",
    "BOXS", "BOXL", "BOXLC", "BOXT",
}

_PREFIXES = sorted(REHAB_TYPES | BASE_TYPES | NONBASE_TYPES, key=len, reverse=True)


def part_type_from_pn(pn):
    """Derive part_type from a part-number prefix (Gen4 MH/RMH/BOX). None if undetermined."""
    if not pn:
        return None
    s = str(pn).strip().upper()
    for p in _PREFIXES:
        if s.startswith(p):
            return p
    return None


# Human part_type words (dispatch / ERP) -> structure_class. Order matters: rehab before base.
# "custom"/"accessory"/etc. carry no class signal -> return None so the part number decides.
NONBASE_WORDS = ("section", "riser", "cone", "lid", "grade", "ring", "transition",
                 "adapter", "troughing", "taper")


def _class_from_word(pt):
    if not pt:
        return None
    s = pt.strip().lower()
    if not s or s in ("none", "(blank)", "nan"):
        return None
    if "rehab" in s:
        return "rehab"
    if "base" in s:
        return "base"
    if any(w in s for w in NONBASE_WORDS):
        return "non-base"
    return None  # custom / accessory / resale / review -> let the part number decide


def coarse_class_from_pn(pn):
    """base / rehab / non-base / None from a part number across all four generations.
    Detection order Gen4 -> Gen2 -> Gen3 -> Gen1 (per item-name-detail-interpreter)."""
    if not pn:
        return None
    s = str(pn).strip().upper()
    if not s or " " in s and not s.startswith(("BOX", "MH", "RMH")):
        return None

    # Gen4: MH / RMH / BOX / MT
    if s.startswith("RMH"):
        return "rehab"
    if s.startswith("BOX"):
        return "non-base" if s.startswith(("BOXS", "BOXL", "BOXT")) else "base"
    if s.startswith("MH"):
        if s.startswith(("MHB", "MHBT")):
            return "base"
        return "non-base"
    if s.startswith("MT"):
        return "non-base"

    # Gen2: [57] + 4 digits + dash + suffix
    m = re.match(r"^[57]\d{4}-(.+)$", s)
    if m:
        suf = m.group(1)
        if suf.startswith(("RS", "SR", "RL", "RC")) or suf.startswith("R"):
            return "rehab"
        if suf.startswith(("BX", "KBX")):
            return "non-base"
        if suf.startswith("B") or suf.startswith("TR"):
            return "base"
        return "non-base"

    # Gen3: [SBCLANFTPMQV] + 4 digits
    if re.match(r"^[SBCLANFTPMQV]\d{4}", s):
        letter = s[0]
        if letter in ("B", "F", "T", "M", "V"):
            return "base"
        if letter == "Q":
            return None  # custom engineered - genuinely ambiguous
        return "non-base"   # S, P, C, L, A, N

    # Gen1: GR / 5GR grade ring, bare rehab/traffic lids, or diameter-leading
    if re.match(r"^5?GR\d", s):
        return "non-base"
    if re.match(r"^(24|27|30)RL$", s):
        return "rehab"
    if re.match(r"^(24|30)TL$", s):
        return "non-base"
    m = re.match(r"^(144|120|96|84|72|60|48)(.*)$", s)
    if m:
        rest = m.group(2)
        tail = re.sub(r"[\d.\-/]", "", rest)   # letters only
        if tail in ("SR", "RS", "CR", "CRC", "RL", "LR"):
            return "rehab"
        if tail.startswith(("B", "BP", "BFF")):
            return "base"
        if rest and re.match(r"^\d+(\.\d+)?$", rest):   # e.g. 6041.33 -> no-letter base
            return "base"
        if tail.startswith(("S", "C", "L", "T", "GR")):
            return "non-base"
        return None
    return None


def structure_class(part_type, pn=None):
    """Classify a part as base / rehab / non-base / unknown.

    The part number is authoritative (it carries the RMH rehab marker that the human
    part_type column often mislabels as 'section'), then the human word, then a Gen4 code."""
    by_pn = coarse_class_from_pn(pn)
    if by_pn:
        return by_pn
    by_word = _class_from_word(part_type)
    if by_word:
        return by_word
    pt_code = (part_type or "").strip().upper()
    if pt_code in REHAB_TYPES:
        return "rehab"
    if pt_code in BASE_TYPES:
        return "base"
    if pt_code in NONBASE_TYPES:
        return "non-base"
    return "unknown"


def year_from_date(v):
    if v is None:
        return None
    if hasattr(v, "year"):
        return v.year
    m = re.search(r"(19|20)\d{2}", str(v))
    return int(m.group(0)) if m else None
