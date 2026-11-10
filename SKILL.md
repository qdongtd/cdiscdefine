---
name: cdisc-define-xml
description: >-
  Author, QC and review CDISC Define-XML 2.1 specification workbooks for SDTM and
  ADaM (the Pinnacle 21 Excel format: Define / Datasets / Variables / ValueLevel /
  Codelists / Dictionaries / Methods / Comments / Documents / Standards / Analysis
  Results) for a clinical study submission. Use whenever the user is drafting,
  filling, syncing, reviewing or fixing a define spec - e.g. "draft the SDTM/ADaM
  define from the spec", "fill origin/source/method/predecessor", "update CRF pages
  from the aCRF", "value level for --ORRES per TESTCD / AVAL per PARAMCD", "fix these
  P21 define or validation findings", "set origins per the CDISC definitions", or a
  cSDRG/ADRG update. Also trigger on P21 rule IDs (DD0035, DD0040, DD0061, DD0073,
  DD0079, DD0129, SD1228...), SDTM/ADaM dataset names (AE, LB, ADSL, ADLB, BDS/OCCDS)
  and files like define_spec.xlsx, *_spec_sdtm.xlsx, *_spec_adam.xlsx, acrf.pdf,
  csdrg.docx, adrg.docx, even when the user doesn't say "Define-XML".
---

# CDISC Define-XML spec builder & QC (SDTM and ADaM)

One skill for both standards. Shared rules are below; the standard-specific
workflows are in [references/sdtm/workflow.md](references/sdtm/workflow.md) and
[references/adam/workflow.md](references/adam/workflow.md) - **read the one that
matches the workbook before filling or reviewing.**

## Which standard?

`scripts/fill_define.py` and `scripts/self_qc.py` detect it from the Datasets tab
(ADaM classes `SUBJECT LEVEL ANALYSIS DATASET` / `BASIC DATA STRUCTURE` /
`OCCURRENCE DATA STRUCTURE` / `ADAM OTHER`, or ADSL/AD* names → ADaM; otherwise SDTM).
Pass `--standard sdtm|adam` to override.

| | SDTM | ADaM |
|---|---|---|
| Fill sources | P21 skeleton + TD SDTM spec (+VLM) + **aCRF** | P21 skeleton + ADaM spec (+ADaM VLM) |
| Origins | Collected (+Source Investigator/Subject/Vendor), Derived, Assigned, Protocol, Predecessor | **Derived, Assigned, Predecessor only** |
| CRF pages | Collected/Investigator\|Subject - actual PDF page | **none** |
| Value level | --ORRES/--STRESC per --TESTCD, SUPP QVAL per QNAM, TS.TSVAL | AVAL/AVALC (etc.) per PARAMCD |
| Workflow | [references/sdtm/workflow.md](references/sdtm/workflow.md) | [references/adam/workflow.md](references/adam/workflow.md) |

```
python scripts/fill_define.py [--standard sdtm|adam] --skeleton <p21_define.xlsx> --spec <spec.xlsx> \
    --out <filled.xlsx> [--vlm <vlm.xlsx>]  [SDTM: --acrf <acrf.pdf> ...]  [ADaM: --issues <issues.xlsx> ...]
python scripts/self_qc.py [--standard sdtm|adam] <filled.xlsx> [SDTM: --acrf <acrf.pdf> --package <folder>]
```

Both are study-agnostic: the study ID, versions, domains, pages and text all come
from the inputs. `--skeleton` is never modified; `--out` is rebuilt each run.

## The one rule that matters most

**Never make anything up.** Every value - origin/source, pages, Predecessor,
Method/Comment text, labels, class, codelist terms, standard/CT versions - comes only
from **guidance** (SDTMIG/ADaMIG, CDISC CT, Define-XML) or the **specs** (TD/ADaM
spec, VLM, P21 define, aCRF, data). No source → leave blank, flag orange, explain in
`Programmer Review`. Copy sourced text **verbatim**. Ask the user (do not assume)
when a choice is theirs: which spec/aCRF version is current, CT version, marker
meanings, whether spec or P21 wins on dataset metadata, how to handle value level
with no VLM column.

## Shared rules (both standards)

- **Origin types** - quote the CDISC Define-XML CT definitions (codelist C170449),
  never paraphrase: Derived "calculated by an algorithm or reproducible rule";
  Assigned "derived through designation, such as values from a look up table or a
  label on a CRF"; Predecessor "a value that is copied from another variable" (not
  limited to another dataset). Derived → Method; Assigned → Comment; Predecessor →
  Predecessor column `DATASET.VARIABLE` (no Source, no Method). A "Set to X.Y" text on
  a Derived/Assigned variable is a Predecessor candidate - report it.
- **Value-level-only (reviewer convention):** once every value-level row of a
  variable has an Origin, leave Origin/Source/Pages/Method/Predecessor/Comment
  **blank on the Variables row**; then drop Methods/Comments nobody references
  (P21 DD0079/DD0080).
- **Spec text priority:** study-specific markers (`[r_def]` ADaM define text,
  `[r_prog]`, `[r_orig]` origin override) › Conversion Definition › Programming Notes.
  Parse each `[r_*]` tag separately; a cell can hold several.
- **Marking:** orange fill = needs human review; red font = a value that overrides
  what was there (spec over P21, or a requested change); routine fills unmarked.
  Add a `Programmer Review` column with the original value + reason. Back up before
  every pass (`_pre_<change>_backup.xlsx`).
- **Dataset metadata** from the spec Content tab (`[r_sort]` override); Key
  Variables always **comma-separated** (P21 DD0040); Datasets `Standard` column filled.
- **Codelists:** unique Names (append the domain/variable); NCI-coded codelists name
  the CT package that holds the code (ADaM CT codes → `ADaM <version>`, else `SDTM
  <version>`, P21 DD0129); terms from the data when available.
- **Wording:** Methods/Comments read as plain English - no librefs, SAS functions,
  placeholders, `[r_*]` markers, spec-layout references or unbalanced quotes. Spec
  text that fails this is copied verbatim and reported for the spec owner to reword.
- **Issue workbook:** every finding goes to an `.xlsx` with its evidence (spec text,
  where a page now points, P21 detail) and a resolution column - never only in chat.
- **Excel-saved workbooks:** openpyxl can corrupt a workbook the user saved in Excel
  (it will not re-open). Always re-open after saving; if it fails, restore the backup
  and apply the edits through Excel (`cscript //nologo` VBScript: `Workbooks.Open`,
  `Cells(r,c).Value`, `Rows(r).Delete` bottom-up, `Save`), then diff before/after.
- **Self-QC is a pre-filter.** Always still generate define.xml and run the P21
  Define validation; triage P21 findings as define metadata / data (programmer punch
  list) / keys (spec owner) / package (files in the zip).

## Layout

```
cdisc-define-xml/
├── SKILL.md                     # this file: routing + shared rules
├── references/
│   ├── sdtm/  workflow.md, origin-source-mapping.md, acrf-page-derivation.md,
│   │          p21-remediation.md, rule-catalog.md, csdrg.md, environment.md
│   └── adam/  workflow.md, adam-rules.md
└── scripts/
    ├── fill_define.py           # entry point -> sdtm_fill.py | adam_fill.py
    ├── self_qc.py               # entry point -> sdtm_qc.py   | adam_qc.py
    ├── common.py                # shared helpers + standard detection
    └── acrf_annotations.py      # aCRF page extraction (SDTM)
```

Requirements: Python 3.12, openpyxl, PyMuPDF (`fitz`, SDTM aCRF), pyreadstat
(data), python-docx (cSDRG/ADRG). Run with `PYTHONUTF8=1`.
