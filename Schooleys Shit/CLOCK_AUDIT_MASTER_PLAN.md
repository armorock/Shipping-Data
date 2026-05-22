# CLOCK Code Audit — Master Plan
**Last Updated:** 2026-05-20  
**Author:** Schooley / Claude  
**Status:** Planning Revised — New Phase 1 Locked In → Ready to Execute

---

## What This Is

A Python-based audit pipeline that cross-references three data sources — BOM, Dispatch Board, and ERP Shipping — to populate the Job Code Master List and surface every meaningful discrepancy. The goal is to answer, for every job code: *what was released to production, what was dispatched, what actually shipped, and where are the gaps?*

---

## The Four Data Sources

| Source | File | Rows | Job Codes | Date Range | ERP |
|---|---|---|---|---|---|
| BOM 2016–2026 | `all_bom_union.xlsx` → `BOM 2016-2026` | 216,005 | 2,104 | 2016–2026 | n/a |
| Dispatch Board | `Dispatch_Board_Master_2019-2025.csv` | 27,353 | 1,924 | 2019–2025 | Manual |
| ERP Shipping | `All Shipping Data BABY.xlsm` → `Master List` | 93,653 | 2,492 | 2014–2026 | QB / FB / NS |
| Job Code Universe | `Job Code Master List.xlsx` → `Master List` | 3,000 | 3,000 | — | — |

---

## Known Data Problems (Catalogued Before Building)

### 1. BOM Revisions
- 16,387 Job Code + Structure Name combinations have more than one BOM revision
- 142 job codes span multiple release years
- No explicit revision number column — revisions identified by BOM Release Date + Source File Name
- **Resolution:** Keep ALL revisions. For each structure, show all revision piece counts. The Dispatch Board + ERP data determines which revision's parts are the "real" assembly.

### 2. Job Code Coverage Gaps
- Only **1,483 job codes** appear in all three data sources
- 230 in BOM → no shipping records anywhere
- 618 have shipping records → no BOM
- 393 in Dispatch Board → no BOM
- 573 in BOM → no Dispatch Board entry
- 464 in Master List → appear in none of the three sources
- 193 in Shipping → not even in the Master List universe

### 3. QB-Era Null Job Codes (~16,500 rows)
- 32% of QuickBooks shipping rows (2014–2023) have no job code
- Job codes didn't exist as a system until ~2016
- Attribution strategy: match by Invoiced Customer (contractor or agency), ship state, and approximate date against BOM records

### 4. No Dispatch Board Before 2019
- Five full years (2014–2018) are ERP-only
- These jobs get the "ERP Only" verification tier; no dispatch confirmation possible

### 5. Structure ID Mismatch
- BOM: "SSMH 10", "SMH D12-004"
- Dispatch Board: "MH 01", "MH 16229039"
- No shared key — requires fuzzy/heuristic matching within job code scope

### 6. Part Number Generation Gap
- Old format: `483C`, `485S`, `484B1.33`
- New format: `MHC4836`, `MHS4848`, `MHB4836.133`
- Bridge: the `NEW NUMBER` column in Shipping (partially populated)
- Secondary bridge: regex-based parsing via the parse-precast-names skill

### 7. Plant Name Inconsistency
- "Sulphur Springs" (QB, NS) vs "Sulfur Springs" (FB)
- Normalize to canonical: **Boulder City (BC)**, **Sulphur Springs (SS)**, **Plant City (PC)**

### 8. Job Code Master List is Empty
- `Date Released to Production`, `BOM Count`, `Shipping Count` are 100% null
- These are all computed fields — this pipeline fills them in

---

## The Audit Logic

The core question per job code is:

```
BOM said → N structures were released to production
Dispatch Board confirmed → M structures were dispatched
ERP Shipping shows → P structures with parts invoiced

If N >> P: FLAG — major shortfall in confirmed shipping
If P > N: FLAG — shipping exceeds BOM (over-shipment or data gap)
If M ≈ P: Healthy — dispatch and ERP align
If M >> P: FLAG — dispatched but no ERP record (shipment not invoiced?)
```

**What counts as a "structure shipped":**  
Every manhole has a top (cone or lid) and a bottom (base or No Bell section). The BOM is the guide for what constitutes the complete assembly for each specific structure. The BOM piece breakdown per structure revision is the reference spec; ERP part counts are compared against it.

### Part Type Classification (all 4 generations)

| Part Type | Gen 1 | Gen 2 | Gen 3 | Gen 4 |
|---|---|---|---|---|
| **Section/Riser** | `484S`, `601SNB` | `5xxxx-s`, `sNB`, `sNB2D` | `S0472-0000`, `S-0001` | `MHS6072`, `MHSxxxxNB` |
| **Base** | `603BP.75`, `482-6BP1.33` | `7xxxx-B75`, `7xxxx-BFF` | `B0436`, `B0548` | `MHB4836.75ES`, `MHBxxxxFF` |
| **Cone (Eccentric)** | `6036C` | `50424-C` | `C0472-000` | `MHC4824` |
| **Cone (Concentric)** | `6036CC` | `50530-CC` | `C0472-001` | `MHCC4836` |
| **Lid** | `484L` | `70612-L`, `70612-HL` | `L0448-00000`, `L-01000` | `MHL96HATCH`, `MHL96CAST` |
| **Transition Lid** | n/a | `50530-TL` | n/a | `MHTL12060`, `MHTLC12060` |
| **Grade Ring** | `GR364`, `GR366` | `50306-GR` | n/a | `MHGR30X6`, `MHGR36X6` |
| **Troughing Insert** | n/a | `7xxxx-TR` | `T-prefix custom` | `MHT`, `MHBT` |
| **Rehab Section** | `484SR` | `5xxxx-Rs` | `S-2000` | `RMH4848`, `RMHC`, `RMHL` |
| **Custom (Job-coded)** | n/a | `sNB3D` variants | `Q/M/F/T/P/V/N prefix` | `BOX`, `BOXS`, `BOXL` |

### Assembly Validation Rules

**Standard structure — requires ALL of:**
1. **1 bottom piece**: a Base (`B`, `MHB`, `7xxxx-B`) OR a No Bell section acting as base (`SNB`, `sNB`, `MHSxxxxNB`, Gen3 `S-0001` as bottom BOM position). The No Bell section as base is identified by: NB suffix present AND it appears first in the structure's BOM part list (lowest height or lowest sequence).
2. **1 top piece**: a Cone (`C`, `MHC`, `MHCC`) OR a Lid (`L`, `MHL`, `MHLC`). A Transition Lid (`MHTL`) is NOT the top — the Cone or Lid above it is still required.
3. **1+ sections**: variable count based on depth.

**Flat Floor structures:** base is a `FF`/`FFES` variant. Assembly logic same — still needs 1 bottom + 1 top + sections.

**Box structures:** `BOX` prefix base + `BOXS` sections + `BOXL` lid. Separate assembly logic.

**Rehab structures:** `RMH` sections + `RMHC`/`RMHL` top. No traditional base — structure wraps existing host.

**Transition Lid rule:** If `MHTL` or `-TL` appears in a BOM, the structure still requires a Cone or Lid above it. A TL-only top is flagged as incomplete in the audit.

**Grade rings and troughing inserts** do not count toward assembly completeness — they are optional accessories.

---

## Revised Phase Structure

The original Phase 1 (normalization) has been pushed to Phase 2. **The new Phase 1 is the Job Code Registry** — establishing the complete job code universe with all known names, locations, and structure lists before any cross-referencing or normalization work begins. Everything else falls into place once we know what job codes exist and what belongs to them.

---

## 8-Phase Execution Plan

---

### Phase 1 — Job Code Registry (NEW STARTING POINT)
**Script:** `01_build_job_code_registry.py`  
**Outputs:**
- `Job_Code_Registry.xlsx` → the Excel master (priority)
- Updated per-job-code markdown files in `/Job Codes/` (append-only, secondary)

**What this builds:**

A single Excel workbook with one row per unique job code, aggregated from all three data sources plus the existing Job Code Master List. This is the foundation document for the entire audit.

**Excel columns:**

| Column | Source | Notes |
|---|---|---|
| `Job Code` | All sources | Trimmed, uppercase |
| `BOM Project Name` | BOM `Project Name` | Most official name |
| `Dispatch Job Name` | Dispatch `job_name` | Often informal — keep raw |
| `Shipping Customer` | QB `Invoiced Customer` | Contractor or agency |
| `State` | BOM first, then dispatch, then shipping | Best available |
| `City` | BOM first, then dispatch, then shipping | Best available |
| `County` | Shipping `Shipping County` | Usually only in ERP |
| `Plant(s)` | Dispatch + Shipping | e.g. "BC, SS" or "BC" |
| `Year Released` | BOM `Year Release` (earliest) | |
| `Date Released to Production` | BOM `BOM Release Date` (earliest) | |
| `In BOM?` | Computed | Y/N |
| `In Dispatch?` | Computed | Y/N |
| `In Shipping?` | Computed | Y/N |
| `BOM Row Count` | Computed | Raw BOM line items |
| `Dispatch Row Count` | Computed | Dispatch board entries |
| `Shipping Row Count` | Computed | ERP line items |
| `Sources` | Computed | e.g. "BOM, Dispatch, Shipping" |

**Markdown file update rules (APPEND ONLY — never clear existing content):**

For each job code with an existing markdown file in `/Job Codes/`:
- Add a `## Structure Registry` section at the bottom if it doesn't exist
- Under `### BOM Structures`: list all unique structure names from the BOM for this job code
- Under `### Dispatch Structures`: list all unique structure IDs from the dispatch board for this job code
- If the section already exists: append any new names not already present, do not remove any existing entries
- Never modify content above the Structure Registry section

For job codes without an existing markdown file: create a minimal new file with the registry data only (no fabricated content).

**Data source priority for location (State, City):**
1. BOM `Job Location` (most complete for 2016+ jobs)
2. Dispatch `city` + `state` (manual but usually accurate)
3. Shipping `Shipping City` + `Shippings State` (ERP — reliable but only has the ship-to address)

---

### Phase 2 — Job Code Validation & Repair (BABY File)
**Script:** `02_jobcode_repair.py`  
**Output:** Annotated BABY file (new columns only, no existing data touched), `jobcode_repair_log.csv`

This phase uses the completed Phase 1 registry as the reference universe to find and fix every bad, partial, or missing job code inside the All Shipping Data BABY file. The BABY file is never destructively edited — all corrections live in new columns.

#### Sub-task A — Partial / Truncated Job Code Repair

Some shipping rows have job codes that are fragments of the real code. Example: `BV` when the correct code is `BVN`. The registry gives us the full list of valid codes, so partial matches can be scored and proposed.

**Detection:** Any job code in the BABY Master List that does NOT exist in the Phase 1 registry is flagged as a candidate for repair.

**Matching logic (in priority order):**
1. **Prefix match:** Find all valid codes that start with the partial value (e.g., `BV` → `BVN`, `BVR`, `BVT`). If only one match exists, it's high-confidence.
2. **Contractor cross-reference:** Match `Invoiced Customer` against BOM `Contractor` and the registry's `Shipping Customer` for the candidate codes.
3. **Location cross-reference:** `Shippings State` + `Shipping City` must match the candidate job code's known location in the registry.
4. **Date window:** `Date Shipped` should fall within ±180 days of the candidate job code's `Date Released to Production`.
5. **Scoring:** Each matching criterion adds points. Score ≥ 3 = high confidence, Score 2 = medium, Score 1 = low, Score 0 = unresolvable.

**Known example:** Job code `BV` in BABY file → likely `BVN`. Large batch. Cross-reference by contractor + city/state that shipped under `BVN` to confirm.

**Output columns added to BABY Master List (new columns only — original `Job Code ` column is NEVER modified):**
- `Fixed_Job_Code` — the single source of truth for all downstream phases. If the original code is already valid, this is a copy of it. If it was repaired, this holds the corrected value. If unresolvable, left blank. **All Phase 3+ scripts read from this column, never from the original.**
- `Repair_Status` — `OK` (original was valid), `REPAIRED` (partial or null code was fixed), `UNRESOLVABLE` (could not match)
- `Repair_Confidence` — High / Medium / Low / n/a (n/a when status is OK)
- `Repair_Method` — which criteria matched (e.g., "Prefix+Location+Contractor"), blank if OK
- `Repair_Notes` — human-readable explanation, blank if OK

> **Safety rule:** The original `Job Code ` column is read-only for all scripts in this pipeline. No script ever writes to it. If a script is broken and goes off the rails, the worst it can do is corrupt the new columns — the original data is always intact and recoverable.

#### Sub-task B — Null Job Code Attribution (QB Era, 2014–2016)

Shipping rows with a completely blank job code. These predate the job code system. Attribution uses the same scoring logic as Sub-task A but without a prefix to anchor on — purely contractor + location + date window matching against known jobs from that era.

**Output columns added (same columns as Sub-task A — one unified set across the whole sheet):**
- `Fixed_Job_Code` — populated with the attributed match if found, blank if not
- `Repair_Status` — `ATTRIBUTED` (null code successfully matched) or `UNRESOLVABLE`
- `Repair_Confidence`, `Repair_Method`, `Repair_Notes` — same as above

---

> ## ⏸ PAUSE & REVIEW CHECKPOINT — Pre-Job-Code Era
>
> **Before proceeding to Phase 3**, both Schooley and Claude review:
>
> 1. The Partial Code Repair log — validate that proposed corrections look right before anything is committed
> 2. The Null Attribution results — determine how many 2014–2016 records were successfully matched vs. remain unresolved
> 3. Decide the structure for permanently unresolvable records: synthetic placeholder codes (e.g., `PRE-2016-[location]`), a single `UNATTRIBUTED` bucket, or leave blank and flag
> 4. Assess the scope of remaining unknowns before proceeding to normalization
>
> **No Phase 3+ work begins until this review is complete and the repair approach is confirmed.**

---

### Phase 3 — Data Normalization
**Script:** `03_normalize_data.py`  
**Output:** Clean CSVs in `/outputs/normalized/`

Tasks:
- Trim + uppercase all job codes across all sources (using validated Phase 1 + Phase 2 registry as reference)
- Normalize plant names → canonical abbreviations (BC, SS, PC)
- Fix "Sulfur/Sulphur" inconsistency throughout
- Parse date fields consistently (handle Excel serials, mixed datetime formats)
- Apply confirmed Phase 2 job code repairs to the working dataset
- Output: `normalized_bom.csv`, `normalized_dispatch.csv`, `normalized_shipping.csv`

---

### Phase 4 — Part Number Crosswalk
**Script:** `04_build_crosswalk.py`  
**Output:** `part_number_crosswalk.csv`

Tasks:
- Extract all old-format PNs from BOM (`Product Number` column)
- Extract all new-format PNs from Shipping (`Part Number` column)
- Use `NEW NUMBER` column in Shipping as the explicit bridge where available
- Use regex parsing (diameter + height + type) to infer mappings where `NEW NUMBER` is null
- Build bidirectional lookup: old ↔ new
- Flag any PNs that can't be mapped

Part type classification logic (by generation, derived from regex on part number):

**Gen 4 (MH prefix):** `MHS`=Section, `MHB`/`MHBT`=Base, `MHC`/`MHCC`=Cone, `MHL`/`MHLC`=Lid, `MHTL`/`MHTLC`=TransitionLid, `MHGR`=GradeRing, `MHT`=Troughing, `RMH`/`RMHC`/`RMHL`=Rehab, `BOX`/`BOXS`/`BOXL`=Box

**Gen 2 (5/7 prefix):** First digit `5`=widget, `7`=base. Suffix: `-s`=Section, `-B`=Base, `-C`=Cone, `-CC`=ConcCone, `-L`=Lid, `-CL`=ConcLid, `-HL`/`-LH`=HatchLid, `-TL`=TransitionLid, `-GR`=GradeRing, `-TR`=Troughing, `-Rs`/`-RL`/`-RC`=Rehab, `-BXs`/`-KBXs`=Box. Suffix ends with `NB` → Nobell modifier.

**Gen 3 (letter prefix):** `S`=Section, `B`=Base, `C`=Cone, `L`=Lid, `Q`/`M`/`F`/`T`/`P`/`V`/`N`=Custom. Option code `0001`=NoBell, `2222`/`2221`/`2122`=LiftersInside, `0020`/`0030`=Doghouse.

**Gen 1 (compact):** Trailing letter `S`=Section, `B`/`BP`=Base, `C`=Cone, `CC`=ConcCone, `L`=Lid, `GR`=GradeRing (prefix). `NB-DH`=NoBell+Doghouse. `SR`/`RS`=Rehab.

**No Bell as Base rule:** Any section with `NB` modifier (`sNB`, `SNB`, `MHSxxxxNB`, Gen3 `S-0001`) that is the first/bottom piece in a BOM is classified as `BASE_NB` rather than `SECTION` for assembly purposes. Detection: lowest sequence position in structure's part list.

---

### Phase 5 — BOM Consolidation by Job Code
**Script:** `05_bom_consolidation.py`  
**Output:** `bom_summary_by_jobcode.csv`, `bom_revisions_detail.csv`

Tasks:
- For each Job Code + Structure Name: group all rows by BOM Release Date
- Each unique BOM Release Date = one "revision"
- Per revision: count total parts, count by type (base, section, cone, lid, grade ring, other)
- Compute: min piece count across revisions, max piece count, latest revision count
- Flag structures where piece count changes significantly between revisions (>2 parts delta)
- Output summary per job code: # structures, total revisions, total parts (latest rev), range (min–max)

Revision hierarchy for "best guess" when no shipping data exists:
1. XML source file types take priority over PDF
2. Most recent BOM Release Date wins within same source type
3. Flag as "Revision Ambiguous" if two XMLs have conflicting counts on same date

---

### Phase 6 — Dispatch Board Confirmation
**Script:** `06_dispatch_matching.py`  
**Output:** `dispatch_confirmed_by_jobcode.csv`, `structure_id_xref.csv`

Tasks:
- For each job code in Dispatch Board: list all structure IDs, ship dates, part numbers
- Attempt structure ID matching to BOM Structure Names:
  - First try: exact match after normalization (remove "MH ", "SSMH " prefixes)
  - Second try: numeric suffix match (both end in same number)
  - Third try: flag as "Unmatched Structure" for manual review
- For matched structures: count confirmed dispatched parts
- Handle 23% null structure_ids: these rows still contribute to job-level part counts even if structure can't be pinned
- Output: per job code → # structures dispatched (matched + unmatched), part count, date range

---

### Phase 7 — ERP Shipping Gap Analysis & Summary
**Script:** `07_erp_attribution.py`  
**Output:** `shipping_summary_by_jobcode.csv`, `qb_attributed_records.csv`, `shipping_unattributed.csv`

**Sub-task 5A — QB Null Job Code Attribution:**
- For each null-job-code QB row: attempt match to a known job code using:
  - `Invoiced Customer` → match against BOM `Contractor` field (both directions: contractor or agency)
  - `Shippings State` → must match BOM `Job Location` state
  - `Date Shipped` → within ±90 days of BOM Release Date for that job code
  - Fuzzy match score: award points for each matching field; threshold = 2+ fields match
- Output: attributed records get a job code + confidence level (High / Medium / Low)
- Remaining nulls → `shipping_unattributed.csv` for manual review

**Sub-task 5B — ERP Summary by Job Code:**
- For each job code: aggregate across QB + FB + NS
- Count total line items, unique part numbers, quantities by part type
- Identify which ERP era the records come from
- Map old part numbers to new format using Phase 2 crosswalk
- Flag: jobs with shipping records only in QB era (no cross-validation available)

---

### Phase 8 — Gap Analysis & Audit Scoring
**Script:** `08_audit_scoring.py`  
**Output:** `audit_flags.csv`, `job_code_audit_summary.csv`

**For each job code, compute:**
- `bom_structure_count` — # unique structures in BOM (all revisions)
- `bom_part_count_latest` — total parts in latest revision
- `bom_part_count_range` — min–max across all revisions
- `dispatch_structure_count` — # structures confirmed on dispatch board
- `dispatch_part_count` — total parts on dispatch records
- `shipping_part_count` — total Armorock parts in ERP for this job code
- `shipping_coverage_pct` — shipping_part_count / bom_part_count_latest × 100
- `dispatch_coverage_pct` — dispatch_structure_count / bom_structure_count × 100

**Flag tiers:**

| Flag | Condition | Priority |
|---|---|---|
| 🔴 CRITICAL | Shipping < 10% of BOM part count | Immediate manual audit |
| 🟠 HIGH | Shipping 10–50% of BOM part count | Review queue |
| 🟡 MEDIUM | Shipping 50–80% of BOM part count | Watch list |
| 🟢 CLEAN | Shipping ≥ 80% of BOM part count | Verified |
| ⚪ NO BOM | Shipping exists, no BOM record | Investigate source |
| ⚫ NO SHIPPING | BOM exists, zero shipping records | Possibly not produced |
| 🔵 ERP ONLY | Pre-2019, no dispatch board coverage | ERP-only verification |
| ❓ UNRESOLVED | QB-era, job code unattributed | Manual attribution needed |

---

### Phase 9 — Output Generation
**Script:** `09_output_generation.py`

**Output A — Job Code Master List (populated)**
- Update `Job Code Master List.xlsx` → `Master List` with:
  - `Year` (from BOM Year Release or earliest ship date)
  - `Date Released to Production` (earliest BOM Release Date for that job code)
  - `Job Name` (from BOM Project Name, filled forward)
  - `State`, `City`, `County` (from BOM Job Location or Shipping address)
  - `BOM Count` (# unique structures in BOM)
  - `Shipping Count` (# unique ERP shipping events)
  - New columns: `Dispatch Count`, `Match %`, `Flag`, `ERP Era`

**Output B — Per Job Code Markdown Files (update existing)**
- For each job code that already has a markdown file in `/Job Codes/`:
  - Append a structured audit block with: BOM structure count, dispatch confirmed, shipping confirmed, flag level, list of unmatched structures
- Create new markdown files for any job codes found only in shipping/BOM with no existing file

**Output C — All Shipping Data BABY (confirm + annotate)**
- Add a `Confirmed_Job_Code` column to the Master List sheet for QB-era rows where attribution was recovered
- Add an `Audit_Flag` column across all rows
- Do NOT modify any existing data — only add new columns

**Output D — Audit Summary Report**
- `AUDIT_SUMMARY.md` — top-level stats: how many job codes Clean / High / Critical, biggest gaps found, QB attribution rate, total structures accounted for

---

## Execution Order & Dependencies

```
Phase 1 (Job Code Registry) ← START HERE
    ↓
Phase 2 (BABY File Job Code Repair + QB Attribution)
    ↓
    ⏸ PAUSE & REVIEW — pre-job-code era decisions
    ↓
Phase 3 (Normalize — uses validated codes from P1+P2)
    ↓
Phase 4 (Crosswalk)   ←── needed for Phase 7
    ↓
Phase 5 (BOM)         Phase 6 (Dispatch)
    ↓                      ↓
           Phase 7 (ERP Gap Analysis)
                   ↓
           Phase 8 (Audit Scoring)
                   ↓
           Phase 9 (Output)
```

Phases 5 and 6 can run in parallel after Phases 1–4 complete.  
The PAUSE after Phase 2 is a hard gate — no Phase 3+ work begins without review sign-off.

---

## Open Questions / Manual Review Items

These need human judgment before or during execution:

1. **Nobel Riser / Transition Lid patterns** — need to confirm exact PN patterns so the "complete assembly" logic handles them correctly
2. **Structure ID naming conventions** — are the numeric dispatch IDs (e.g., "MH 16229039") from a specific ERP field, or hand-typed? This affects fuzzy matching confidence
3. **QB era invoiced customer names** — run a sample match first and review confidence before mass-attributing
4. **Structures with zero parts in every revision** — are these placeholder rows or legitimate structures that just use non-Armorock hardware?
5. **Job codes in Shipping not in Master List (193)** — these will be added to the universe but need a source-of-truth review

---

## File Structure (outputs)

```
FULL JOBCODE SHIPPING HISTORY/
├── MASTER CSV FILES/          ← source files (read-only)
├── outputs/
│   ├── normalized/
│   │   ├── normalized_bom.csv
│   │   ├── normalized_dispatch.csv
│   │   ├── normalized_shipping.csv
│   │   └── job_code_universe.csv
│   ├── crosswalk/
│   │   └── part_number_crosswalk.csv
│   ├── bom/
│   │   ├── bom_summary_by_jobcode.csv
│   │   └── bom_revisions_detail.csv
│   ├── dispatch/
│   │   ├── dispatch_confirmed_by_jobcode.csv
│   │   └── structure_id_xref.csv
│   ├── shipping/
│   │   ├── shipping_summary_by_jobcode.csv
│   │   ├── qb_attributed_records.csv
│   │   └── shipping_unattributed.csv
│   └── audit/
│       ├── audit_flags.csv
│       ├── job_code_audit_summary.csv
│       └── AUDIT_SUMMARY.md
├── CLOCK_AUDIT_MASTER_PLAN.md  ← this file
└── Job Codes/                  ← existing per-job-code markdown files
```

---

## What to Build First

**Start with Phase 1 — the Job Code Registry.** This is the right first step because it answers "what job codes exist and what do we know about them?" before touching any audit logic. The Excel output becomes the working document you can visually validate immediately.

Then **Phase 2 (normalization)** cleans the raw data, followed by **Phase 3 (crosswalk)** which bridges old and new part number formats.

**Phase 6 QB attribution** is the highest-risk phase and should be reviewed manually before any mass-write back to the BABY file.

## Tool to Use for Execution

**Claude Code (VS Code)** is the better tool for running the actual Python scripts — it has persistent terminal sessions, can handle large file reads iteratively, and lets you watch script output in real time. Claude Cowork is better for planning, reviewing outputs, and making targeted edits.

Recommended workflow: Plan and design in Cowork → Execute scripts in Claude Code → Review outputs back in Cowork.
