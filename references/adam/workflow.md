# ADaM Define-XML workflow (full detail)

ADaM part of the cdisc-define-xml skill; same working conventions as SDTM, ADaM
rules. ADaM has **no CRF pages and no Collected/Protocol origins** — every variable
is **Derived, Assigned or Predecessor**.

## The one rule that matters most

**Never make anything up.** Origins, Predecessor variables, Method/Comment text,
dataset class/structure/keys, codelist terms and standard/CT versions come only from
**guidance** (ADaMIG, CDISC CT, Define-XML) or the **specs** (the ADaM spec, the
ADaM VLM, the P21-generated define, the ADaM data). No source → leave blank, flag
orange, explain in `Programmer Review`. Copy spec text **verbatim**.

## Inputs

1. **P21-generated ADaM define** from the study's ADaM `.xpt` — the skeleton
   (datasets, variables, value-level rows per PARAMCD, codelists from the data).
   P21 guesses many origins as Predecessor; the spec decides.
2. **Compound ADaM spec** (TechData template: Study, Content, one tab per dataset,
   study columns, parameter tabs such as `LBPARAMN`) **+ ADaM VLM** (one tab per
   dataset: VariableName / WhereVarList / Comparator / WhereValue / Origin /
   Conversion Definition / Programming Notes / study columns).
3. *(optional)* the ADaM data, to derive codelist terms (never trust spec codelists
   over the data — see references/adam-rules.md).

## Run the fill script

```
python scripts/fill_define.py --skeleton <p21_adam_define.xlsx> --spec <adam_spec.xlsx> \
    --out <filled.xlsx> [--vlm <adam_vlm.xlsx>] [--issues <issues.xlsx>] [--study <ID>] \
    [--ct <YYYY-MM-DD>] [--keep-p21-dataset-meta]
python scripts/self_qc.py <filled.xlsx>
```

Study-agnostic: the study ID comes from the Define tab, IG/CT versions from the
spec Study tab, scope from the Content-tab study column. It never modifies
`--skeleton`. **Red font** = value that overrides the P21 skeleton (spec wins);
**orange** = needs review; every changed/uncertain cell has a Programmer Review note.
`--issues` writes the issue workbook (Summary + Issue Details, one row per finding
with the spec text as evidence; repeated value-level findings grouped per variable).

**Ask the user (do not assume)** when it is unclear: which spec/VLM version is
current (several copies often exist — prefer the ones in the study folder), what
an unfamiliar study-column marker means, and whether the Content tab or P21 wins on
dataset Class/Structure/Key Variables (`--keep-p21-dataset-meta` keeps P21 and only
reports).

## Reading the ADaM spec

- Study column markers (parse each tag separately; a cell can hold several):
  `x` = in scope; `[r_def]` = study-specific **define text**; `[r_prog]` =
  study-specific derivation; `[r_orig] <Origin>` = study origin override;
  `[r_sort]` (Content tab) = study key/sort order.
- **Method/Comment text priority:** `[r_def]` › `[r_prog]` › Conversion
  Definition › Programming Notes.
- `<COREVAR>` row in a dataset tab = the ADSL core variables (ADSL tab `Core
  Variable = Yes`) are carried into that dataset → **Predecessor `ADSL.<VAR>`**, with
  type/length from the ADSL row.
- Origin by spec: **Derived → Method** `MT.<DS>.<VAR>` (Source Sponsor);
  **Assigned → Comment** `COM.<DS>.<VAR>` (Source Sponsor); **Predecessor → the
  Predecessor column** `DATASET.VARIABLE` (no Source, no Method). Value level:
  `MT.<DS>.<VAR>.<PARAMCD>`.
- Predecessor text must be a plain copy: `DM.AGE`, `Set to ADSL.TRTP`, or
  `SUPPAE.QVAL where QNAM="X"`. When the spec says Predecessor but the text carries a
  derivation (`LC.LCSTRESN. If missing…`, `PC.PCSTRESN/1000`), keep the source
  variable, flag orange and report "consider Derived". When the spec says
  Derived/Assigned but the text is a pure copy, report it as a Predecessor candidate
  (CT C170550 "a value that is copied from another variable") — don't change it.

## Value level (PARAMCD)

Match each skeleton value-level row to the VLM row of the same dataset, variable and
where clause (comparison ignores quotes, parentheses, spacing and case). VLM Origin
blank → the parent variable's spec origin. **No VLM row → use the parent variable's
spec row** (sourced; reported once per variable with its PARAMCD list, so the
programmer can check the parent text fits each parameter). VLM rows marked for the
study that match no row in the data are reported (often AVAL in the VLM where the
data only has AVALC, or a CHGCATx the data doesn't carry).

**Reviewer convention:** once every value-level row of a variable has an Origin,
leave Origin/Source/Method/Predecessor/Comment **blank on the Variables row** (model:
ADEG AVAL/AVALC). Checked by `self_qc.py`.

## Dataset metadata, Standards, Codelists

- Class/Structure/Key Variables from the **Content tab** (`[r_sort]` override)
  unless told otherwise. Valid Define-XML 2.1 ADaM classes: `SUBJECT LEVEL ANALYSIS
  DATASET`, `BASIC DATA STRUCTURE`, `OCCURRENCE DATA STRUCTURE`, `ADAM OTHER`. P21
  can emit `SYSTEM` for a sponsor dataset — replace it from the Content tab. A BDS
  dataset without PARAMCD/PARAM is reported (Content-tab class may be wrong).
- Datasets `Standard` = `ADaMIG <version>`; Standards tab = ADaMIG + CT rows from
  the spec Study tab (flagged to confirm the ADaM vs SDTM CT release).
- Blank DataType/Length in the skeleton (no-data variables) → spec Type/Length;
  date types carry an empty Length.
- Duplicate codelist Names → append `(<DS>.<VAR>)` for value-level codelists.
- Codelist terms: derive from the ADaM data when available (`pyreadstat`,
  `disable_datetime_conversion=True`), never from a spec's 1..n assumption.

## Editing a workbook the user has saved in Excel

A define spec the user has opened and saved in Excel can carry a stylesheet that
openpyxl corrupts on resave (the saved file no longer re-opens:
`IndexError` in `_expand_named_style`). Always re-open after saving; if it fails,
**restore the backup** and apply the edits through Excel itself (a VBScript run with
`cscript //nologo`: `Workbooks.Open`, `Cells(r,c).ClearContents` / `.Value`,
`Rows(r).Delete` bottom-up, `Save`). Plan the edits with openpyxl read-only first,
then verify with a before/after diff that nothing else moved.

## P21 data-validation triage (ADaM)

Classify each open issue: **define metadata** (fix in the workbook), **data** (ADaM
programming punch list), **key variables / data** (spec owner decides), **package**
(files in the zip). Common ones:
- **SD1228** value not in a user-defined codelist (value-level AVALC codelists built
  from another data cut): remove the value-level codelist (unlink + delete when no
  longer referenced) or rebuild it from the current data - ask which.
- **SD1152** duplicate records per Key Variables: show the duplicated key values; null
  key values are a data issue, a key that cannot separate records needs another key
  variable (e.g. a --SEQ) - do not change keys without the spec owner.
- **DD0040** missing Key Variables: often a space-separated key - make it
  comma-separated (same variables).
- **DD0084 / DD0085**: referenced document (ADRG PDF) / define2-1-0.xsl missing from
  the package.
- Removing a dataset: delete its Datasets/Variables/ValueLevel rows, then any
  Method/Comment/Codelist no longer referenced anywhere.
Write the triage (issue, source, resolution, P21 evidence) to a workbook.

## Self-QC

`self_qc.py` is a pre-filter (still run P21): ADaM origin set; Derived needs
Method; Method only on Derived; Predecessor value + `DATASET.VARIABLE` syntax;
Source rules; no Pages; origin present (variable or value level); value-level-only
convention; DD0074; data type/length; BDS has PARAMCD/PARAM; dangling IDs; duplicate
codelist names; referenced codelists exist; reviewer-readable wording (no librefs,
SAS functions, placeholders, `[r_*]` markers, unbalanced quotes); copy text not
Predecessor.

More detail and rationale: [references/adam-rules.md](adam-rules.md).

## A fresh P21 export after the define was filled

When the user re-exports from P21 (new data cut or new datasets), the export comes
back with empty Methods/Comments and P21's guessed origins everywhere. Check for this
first (Methods/Comments tabs empty) and ask before filling. The usual answer is
**carry forward + refill**:

1. Back up the export (`_p21export_<date>_backup.xlsx`) - it is the new skeleton.
2. For every Variables / ValueLevel row that exists in the last filled define, copy
   Origin, Source, Pages, Method, Predecessor, Comment (and the Method/Comment texts)
   from it - this keeps the user's manual fixes and earlier P21-triage decisions.
3. Copy Data Type / Length / Significant Digits only where the export is blank
   (no-data variables); otherwise keep the export's values - they reflect the new data
   (e.g. AVAL length/significant digits can change between data cuts).
4. Rows that are new since the last define: fill from the spec (run the fill on the
   export and take those rows), and list each one in the issue workbook.
5. Re-apply earlier define decisions the export undoes (e.g. user-defined codelists
   removed after P21 SD1228 come back), carry Datasets metadata, Standards, codelist
   Terminology/Names, then de-duplicate codelist Names again - new datasets often bring
   codelists whose names collide with existing ones (DD0131).
6. Verify: carried rows equal the last define on every origin field; only data
   attributes may differ.

## Datasets specified outside the compound spec

Some datasets come with their own requirements workbook (e.g. an NCA vendor's
ADNCA/ADPP specs: Variable Name, Type, Origin, Source/Derivation, Comment). Use it as
the spec for those datasets:
- Predecessor: `Source/Derivation` when it is `DATASET.VARIABLE`; a bare variable
  name is a copy within the same dataset (`ADNCA.AVISIT`, flag it); two sources
  (`PC.VISIT, EX.VISIT`) are kept verbatim and flagged - confirm or make Derived.
- Derived / Assigned: Method / Comment text = Source/Derivation and the Comment
  column verbatim, on separate lines (ask the user if they prefer another form).
- Value level (AVAL per PARAMCD): the parent variable's requirements row.
- Variables in the data but not in the requirements (typically ADSL variables): keep
  P21's Predecessor and report them once.
- Blank data types: same variable elsewhere in the define; else the requirements
  `Type` (Char -> text) with Length left blank and flagged.
- Dataset Class/Structure/Key Variables are usually not in such workbooks - keep
  P21's (comma-separated keys) and flag for confirmation.
