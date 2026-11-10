# cdisc-define-xml

A Claude Code skill for drafting, QC-ing and reviewing CDISC **Define-XML 2.1**
specification workbooks (Pinnacle 21 Excel format) for **SDTM and ADaM**.
It replaces the separate `sdtm-define-xml` and `adam-define-xml` skills.

```
python scripts/fill_define.py [--standard sdtm|adam] --skeleton <p21_define.xlsx> --spec <spec.xlsx> --out <filled.xlsx> ...
python scripts/self_qc.py     [--standard sdtm|adam] <filled.xlsx> ...
```

The standard is detected from the workbook's dataset classes. SDTM uses the TD SDTM
spec, VLM and annotated CRF (actual PDF page numbers); ADaM uses the ADaM spec and
ADaM VLM (Derived / Assigned / Predecessor, value level per PARAMCD).

Shared rules: never fabricate (blank + flag instead), CDISC Origin Type definitions,
value-level-only origin convention, plain-English Method/Comment wording, issue
workbooks with evidence, safe editing of Excel-saved workbooks.

See `SKILL.md`, `references/sdtm/workflow.md` and `references/adam/workflow.md`.

Requirements: Python 3.12, openpyxl, PyMuPDF, pyreadstat, python-docx.
