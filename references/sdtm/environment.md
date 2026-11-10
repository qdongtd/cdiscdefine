# Environment, tooling, and workbook layout

## Toolchain
- **xlsx**: `openpyxl`. Read header row once, build a name→index map:
  `h=[c.value for c in ws[1]]; ci={n:i for i,n in enumerate(h)}`, then address
  cells as `row[ci['Variable']]`. Do **not** use `ws.columns` — some P21
  workbooks are missing `<tblGrid>` and `.columns` raises; iterate
  `ws.iter_rows()` / `ws[1]` instead.
- **PDF (aCRF)**: PyMuPDF (`import fitz`). `page.get_text('dict')` for spans,
  `page.get_drawings()` for box rectangles + fill colors, `doc.get_toc()` for
  bookmarks. Don't trust a stated page count — open it and check `doc.page_count`.
- **XPT data**: `pyreadstat.read_xport(path, encoding='latin1',
  disable_datetime_conversion=True)`. The datetime flag keeps ISO 8601 `--DTC`
  strings as text instead of mangling partial dates.
- **docx (cSDRG)**: `python-docx`. To replace a paragraph's text while keeping
  its formatting, set `p.runs[0].text = new` and blank the rest:
  `for r in p.runs[1:]: r.text=''`.
- Always run with `PYTHONUTF8=1` (Windows console is cp1252 and clinical text
  carries en-dashes, µ, degree signs, smart quotes).

## Define workbook tabs and key columns
Column names vary slightly by P21 template version — always resolve by header
name, never by hard-coded letter. Common columns:

- **Datasets**: `Dataset`, `Description`/`Label`, `Class`, `Structure`,
  `Purpose`, `Repeating`, `Reference Data`, `Comment`, `Has No Data`, `Pages`.
- **Variables**: `Order`, `Dataset`, `Variable`, `Label`, `Data Type`, `Length`,
  `Significant Digits`, `Format`, `Mandatory`, `Codelist`, `Origin`, `Source`,
  `Pages`, `Method`, `Comment`, `Role`, `Has No Data`.
- **ValueLevel**: `Order`, `Dataset`, `Variable`, `Where Clause`, `Data Type`,
  `Origin`, `Source`, `Pages`, `Method`, `Comment`, `Mandatory`. The
  `Where Clause` string encodes the qualifier, e.g. `QNAM EQ AETRTEM`.
- **Codelists**: `ID`, `Name`, `NCI Codelist Code`, `Data Type`, `Terminology`,
  `Order`, `Term`, `NCI Term Code`, `Decoded Value`.
- **Methods**: `ID`, `Name`, `Type` (usually `Computation`), `Description`,
  `Expression Context`, `Expression Code`.
- **Comments**: `ID`, `Description`, `Document`, `Pages`.
- **Standards**: the CT/IG versions. When the user asks to bump CT (e.g. to
  `SDTM 2025-09-26`), update the CT row's version/label here and any TS
  `TSVCDVER` references that must agree.

## ID conventions
- Method IDs: `MT.<DATASET>.<VAR>` (or `MT.SUPP<DS>.<QNAM>` for value-level).
- Comment IDs: `COM.<DATASET>.<VAR>`.
- Keep a shared `CM.NODATA` comment for `Has No Data=Yes` variables — when you
  reassign origins later, preserve `CM.NODATA` links; only unlink `COM.*`
  comments you are replacing.

## SUPP-- structural variables (never need method/comment)
`{STUDYID, RDOMAIN, USUBJID, IDVAR, IDVARVAL, QNAM, QLABEL, QVAL, QORIG, QEVAL}`
are the SUPP shell — exclude them from "missing method/comment" QC sweeps.

## Highlight / backup snippet
```python
import openpyxl, shutil
from openpyxl.styles import PatternFill
ORG = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')  # needs human review
shutil.copy(F, F.replace('.xlsx', '_pre_edit_backup.xlsx'))
wb = openpyxl.load_workbook(F)
# ... edits: routine filled-from-source cells get NO fill; ORG only on cells a human must review ...
try:
    wb.save(F)
except PermissionError:
    F = F.replace('.xlsx', '_updated.xlsx'); wb.save(F)
print('saved', F)
```

## Orange-only highlight + Programmer Review column
- **Orange** (`FFC000`) = a cell that needs human review (unsourced value,
  spec/aCRF or skeleton/spec conflict, missing codelist, unresolved item). Do
  **not** highlight routine fills — track changes via the `_pre_*_backup.xlsx` diff.
- Add a non-standard **`Programmer Review`** column to Datasets/Variables/ValueLevel
  (P21 ignores unknown columns). For any **overridden** (aCRF/data over spec) or
  **uncertain/conflicting** value, keep the accepted value and record the
  **original value + reason** there — e.g. `Pages: spec=99, aCRF=85 -> used aCRF` or
  `Method: N/O differ - N used; O="…"`. Never silently change; the programmer
  reconciles these before finalizing.
