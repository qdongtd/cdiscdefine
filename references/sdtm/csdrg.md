# cSDRG (clinical Study Data Reviewer's Guide) updates

The cSDRG is a Word document (`csdrg.docx`) that accompanies the define. Edit with
`python-docx`, keeping formatting intact.

## Formatting-safe paragraph/cell edit
```python
def setp(p, text):
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ''
    else:
        p.add_run(text)

def set_cell(cell, text):
    setp(cell.paragraphs[0], text)
    for extra in cell.paragraphs[1:]:
        for r in extra.runs:
            r.text = ''
```
Back up first (`shutil.copy('csdrg.docx','csdrg_backup.docx')`). Domain
description text and conformance explanations live in tables; walk
`d.tables[i].rows` and match on the first cell (the domain code or rule id).

## What to keep consistent with the define
- **Domain labels must match the define's Datasets tab.** If you rename a domain
  label in the define (e.g. `LC` → `Laboratory Findings in Conventional Units`),
  update every place it appears in the cSDRG: the domain summary table, the
  section-3.4.x heading, and the Table of Contents entry (leave the TOC "right
  click > Update Field" note if the template has one).

## Conformance-issue explanations
Section that lists remaining P21 conformance findings needs a plain-language
explanation per finding. Source each explanation from the **actual** issue (the
punch-list you built), not a generic template. Explain *why* the finding is
acceptable or how it was resolved (e.g. "Codelist is extensible; value reported
as collected", "smart apostrophe replaced with ASCII apostrophe in final TI").

## Narrative sections from a reference RG
When a sibling study in the same program already has a completed reviewer's guide,
use it as the template for narrative sections (trial design derived from TS
parameters, per-domain descriptions, RELREC relationships, coding
dictionary/versions). Copy the *structure and factual pattern*, but substitute
this study's real values (study id, arm count, enrollment, MedDRA/WHODrug
versions, NCT number). Do not carry over the reference study's numbers.

## Items that usually need the user
Appendix of inclusion/exclusion criteria and the issues-summary section often
need study-specific text the user must confirm — fill what you can source and
flag the rest.
