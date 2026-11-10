# Pinnacle 21 Define-issue remediation

Work from the **downloaded Define issue report** (CSV). **Triage first:**

- **Define-metadata** — the workbook disagrees with itself or the standard. Fix in
  the Excel spec (recipes below).
- **Data / programming** — the actual SDTM `.xpt` data is wrong. **Not** a define
  fix → put it on the punch-list deliverable (format at the bottom) and leave the
  define alone.

Always confirm the rule's meaning against the specific finding text before
applying a recipe — the same rule id fires for different reasons. When a fix needs
a clinical/data judgement you can't source, leave it and flag it (Programmer
Review / punch-list) rather than guessing.

Before generating define.xml, run [`../scripts/self_qc.py`](../scripts/self_qc.py)
to catch the spec-checkable rules and cut round-trips. It is a **pre-filter, not
authoritative** — ~220 of the ~279 catalog rules need the generated define.xml /
CT DB / MedDRA / standard-label DB / data, so still run the real P21 Define
validation. When self-checking, remember to scan **both** the Variables and the
**ValueLevel** tabs (a Variables-only check silently passed value-level `eDT`,
missing methods, and origin mismatches on this pilot).

## Define-metadata recipes

| Rule | Meaning | Fix |
|---|---|---|
| **DD0035** | Missing Pages (Collected + Investigator/Subject) | Page from the base Pages column / aCRF (see acrf-page-derivation). If the var is vendor-assay data → Source `Vendor` (no page). |
| **DD0042** | Derived variable missing Method | Add `MT.<DS>.<VAR>` from the spec text (priority `[r_prog]›N›O`) or the METHODS tab (e.g. standard VISITNUM/EPOCH algorithm). **Value-level too** — a Derived value-level slice (e.g. a QS score) needs a **Method**, not a Comment; move the text if it was put in a Comment. |
| **DD0113** | Missing MedDRA reference on coded var | `--DECOD/--LLT/--PTCD/--HLT/--SOC/…` MedDRA-coded vars must reference the MedDRA dictionary (e.g. `PR.PRDECOD → MedDRA`). Note `DSDECOD` is disposition-coded, not MedDRA. |
| **DD0055** | Invalid dataset Class | Set to a valid GNRLOBSC value (e.g. custom `DC` `SYSTEM` → `SPECIAL PURPOSE`, like DM). |
| **DD0068** | Invalid use of Length | Length must be **empty** when Data Type is not integer/float/text → **clear** the length on date/time types (e.g. `BRTHDTC` partialDate). |
| **DD0073** | Invalid Origin Type value | `eDT` is invalid in Define-XML 2.1 → `Collected` / Source `Vendor`. **Check the ValueLevel tab too** — SUPP QVAL rows often still carry `eDT`. |
| **DD0074** | Variable vs value-level Origin mismatch | The variable Origin must match *every* value-level Origin. If they all match, align the variable to it; **if the values have differing origins** (some Assigned, some Derived — e.g. `AU.AUORRES`, `QS.QSSTRESC`), **clear the variable-level Origin** so origin is defined at the value level only. |
| **DD0013** | Invalid integer value for Length | A non-integer `Length` (often **page numbers leaked into the Length column** — e.g. `Length="80 82"`) → set a real integer length; the pages belong in the `Pages` column. |
| **DD0075 / OD0075** | Invalid / null Data Type | Assign: `--DTC`→partialDatetime (BRTHDTC→partialDate), `--DUR`→durationDatetime, `--DY`→integer, else `text`. |
| **OD0080** | Codelist/Variable Data Type mismatch | Set the variable Data Type to match its codelist's type (usually `text`). |
| **DD0108** | STUDYID Origin not Protocol | Set `STUDYID` → `Protocol`/Sponsor. |
| **DD0110** | Inconsistent Origin for VISITNUM across datasets | Make VISITNUM consistent (usually `Derived`) everywhere; add the VISITNUM method where newly Derived. |
| **DD0123** | Variable/value-level Length mismatch | Set the value-level Length to the variable Length (value-level ≤ variable). |
| **DD0124** | Expected codelist missing | Link the CT codelist. Coded vars point to their dictionary (e.g. `PRDECOD → MedDRA`, `CMDECOD → WHODD`). Missing/un-sourceable → `DATASET.VARIABLE` name + orange highlight + Programmer Review. |
| **DD0127** | NCI code without Terminology / codelist collision | If it's a real CDISC codelist → add the **Terminology version** (make it standard). If a sponsor subset collides with a CT name → keep it sponsor and ensure **no** NCI code. |
| **DD0131** | Duplicate codelist Name | Make each Name unique — append a context suffix: `Laboratory Test Name (LB)` vs `(LC)`, `Dictionary Version (SUPPAE)`. |
| **DD0134** | HasNoData variable needs a comment | These are `Has No Data = Yes` vars → link a shared `COM.NODATA` ("planned but no data were collected…"). |
| **DD0136** | Invalid dataset Label | Set to the CDISC standard label (`TS`→"Trial Summary"; a split lab domain `LC`→"Laboratory Test Results"). |
| **DD0137** | Variable label differs from CDISC standard | Usually a sponsor decision → defer to the programmer / note in the cSDRG conformance table. |
| **DD0148** | Invalid Predecessor syntax | The Predecessor must be `DATASET.VARIABLE` (or `DATASET.VARIABLE where VAR="VALUE"`). Move any extra prose into the Comment. |
| **DD0150 / DD0021 / DD0138** | Standards declaration invalid / missing | See the Standards fix below. |
| **OD0070** | Missing Length (integer/float/text) | Assign a length (identifiers from convention: QEVAL 22, IDVAR 8, IDVARVAL 200; else from the spec Length column; else int→8/text→200). |
| **OD0071** | Missing Significant Digits (float) | Set a value (placeholder → flag in Programmer Review to confirm precision vs SAP). |
| **OD0048** | Dangling Method/Comment reference | Add the missing Methods/Comments row, or drop the reference. |
| **CT200x** | Value not in a codelist | Extensible codelist → allowed extension, note it. Not extensible → data issue (punch-list). |

## Standards fix (DD0150 / DD0021 / DD0138)

The CT standard must be declared **and** every codelist's Terminology value must
reconcile with it.

1. **Standards tab** (add if missing) with two rows:
   - `SDTMIG` · Type `IG` · Version (e.g. 3.4) · Status Final
   - CT · Type `CT` · **PublishingSet `SDTM`** · Version (e.g. 2025-09-26) · Status Final
2. **Codelist Terminology** on every NCI-coded codelist must parse to the **same
   Name + Version** as the Standards CT row. P21 derives a Standard from the
   Terminology string by splitting off the trailing date as the Version and the
   rest as the Name — so the Name part must be a **valid STDNAM** and must match
   the Standards row:
   - `CDISC/NCI 2025-09-26` (Terminology) ↔ Standards Name `CDISC/NCI`
   - or `SDTM 2025-09-26` ↔ Standards Name `SDTM` (program convention)
   - **not** `CDISC SDTM CT 2025-09-26` → Name `CDISC SDTM CT` is invalid (DD0021).
3. Per the SOP, these are pre-filled by P21 (Input 1) — **verify** them; complete
   only if blank. Confirm the CT version with the user; it follows the TD spec.

## Data / programming punch-list (deliverable — don't edit the define)

Produce each round and hand to the SDTM programmer. Three sections:

- **Errors (must fix):** `# | Dataset | Rule | #Records | Problem | Suggested fix`.
- **Warnings / Notices (review, not blocking).**
- **Left per instruction:** deferred items the programmer said they'd check
  (e.g. DD0137 labels, SD1152 duplicate keys) — tracked so they don't get lost.

Typical data findings: smart-quote / non-ASCII in `IETEST` (SD1029/SD0037), `.0`
float artifact in integer `QVAL` (SD1228/SD1230), `--TESTCD`↔`--TEST` CT mismatch
(CT2003), missing companion var in data (SD1283), ongoing use missing
end/relative-time (SD0021), value↔code mismatch vs MED-RT (SD2263/SD2265),
all-null permissible vars (SD1078 — acceptable if `Has No Data=Yes`).
