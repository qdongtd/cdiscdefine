# Deriving CRF page numbers from the annotated CRF (aCRF)

**Page source priority:** the study column's `[r_page] N` override → the domain
tab's **base Pages** column (for `x`-marked variables) → **derive from the aCRF**.
Required for Collected + Investigator/Subject (2.1) / CRF (2.0); vendor data needs
no page.

The aCRF's **annotation layer** is authoritative for page **values**. Use
[../scripts/acrf_annotations.py](../scripts/acrf_annotations.py) — it already
implements everything below. Read this file to understand *why* each filter exists
so you can debug an off page number.

## Cover-page offset — printed page ≠ PDF page (do this FIRST)

**Team rule: define Pages are actual PDF page numbers** (what the define.xml link
opens). The extractor's `offset` argument selects the numbering: `offset=0` →
PDF page (1-based) — use this for the define; `offset=detect_cover_offset(doc)` →
printed footer page. The cover count is still needed to translate a TD spec whose
Pages follow the printed footer (`fill_define.py` detects and shifts automatically).

A PDF page index is **not** the printed CRF page number: aCRFs have front matter
(a generated-on cover, a TOC). If you number pages by PDF index you will be off by
a constant and then "correct" every page **wrongly**. `detect_cover_offset(doc)`
finds the PDF page that prints `1 of N` and returns the front-matter count; then
`printed = pdf_0based − offset + 1`. `extract(pdf, offset=None)` applies it
automatically (`--cover-offset auto`); pass an int to override. Verify once: pick a
clean variable (e.g. `CMTRT`) and confirm its extracted page equals the spec page
in the same numbering. A single generated-on cover page is common → cover count 1,
so PDF page = printed page + 1.

## What an annotation is
Annotations are the SDTM variable names overlaid on the blank CRF. They are set
in **`ArialMT` / `Arial-BoldMT`** font; the CRF form body itself is a different
font (often `Courier`). So the first filter is: keep only Arial-family spans.

## The four filters (in order)

1. **NOTE-box exclusion.** Text sitting in an annotation box whose content
   starts with `note:` / `note -` (case-insensitive) is explanatory prose, not a
   variable annotation. Drop the whole box. A variable that appears *only* inside
   a NOTE box on a page does **not** get that page.

2. **Structured-span filter.** A real annotation is a variable token or a short
   structured expression (`AESToken`, `X when Y = Z`, `DOMAIN.VAR`). Reject spans
   that read like sentences — i.e. contain lowercase words other than the
   connectors `in | when | and | or | eq | ne | to | per | of`.

3. **RELREC handling — count the LEFT var only.** A `RELREC when X = Y`
   annotation relates two variables; it should count the page only for the
   left-hand variable of the relationship, not both. Without this, a variable
   named on the right side of many RELREC notes (e.g. `AESPID`, `AELNKID`) gets
   dozens of spurious pages.

4. **Color → domain disambiguation.** Each domain's annotations sit in a box with
   a **distinct fill color**; the annotation *text* is always black, so color is
   the only domain signal. A page can carry boxes for several domains. Assign a
   token to a page for variable `V` only when the box it sits in belongs to one of
   `V`'s domains.

   - Map color → domain from the box that contains a domain header span matching
     `^([A-Z]{2})\s*\(` (e.g. `AE (Adverse Events)`).
   - For each candidate token span, find the **smallest** box that contains it
     (`(rect & span_rect).get_area() > 0.3 * span_rect.get_area()`), and take that
     box's fill color as the token's domain color.
   - **Round fill colors to 1 decimal** before comparing — the same visual color
     can differ in the 2nd/3rd decimal between the header box and a field box
     (e.g. `0.59` vs `0.60`) and an exact match will miss.
   - Resolve `V`'s domain(s) from the define itself (`var2domains`). Treat
     `DM`/`DC` as equivalent (demographics as-collected shares DM's palette), and
     be careful with 2-letter variable stems that collide with domain codes
     (`AGE`→AG, `SEX`→SE, `DTHDTC`→DT) — match on the box color, not on slicing
     the variable name.

   - **Assign pages per DATASET ROW, not per variable.** A name like `FOCID`,
     `--SPID` or a generic identifier lives in **several** domains; if you flatten
     to `{var: pages}` and drop it onto every row of that variable, `AE.FOCID`
     inherits the OE/MH/SC pages too. Use `collapse_by_domain(var_pages)` →
     `{(var, domain): pages}` and fill a Variables-tab row from the entry matching
     **that row's own dataset domain** (`AE.FOCID` → only the `AE`-colored pages).
     Keep the extractor's `(page, domain)` pairs; only fall back to the
     domain-blank (`''`) entries as a name-scan when the row's domain has none.

## Extractor pitfalls (fixed in `acrf_annotations.py` — keep them fixed)
- **Exclude the NOTE box, never its color.** A NOTE box usually shares its domain's
  fill color; excluding by color drops every genuine annotation of that domain on the
  page (e.g. all AE seriousness fields on a page that also carries an AE NOTE).
- **Continuation pages have no domain header.** A multi-page form shows the
  `XX (Domain)` header box only on its first page; map a box color to a domain
  PDF-wide when that color belongs to exactly one domain, and reuse it on header-less pages.
- **Tokens glued to `/` or `=`.** `DSTERM/DSDECOD=INFORMED CONSENT`, `PRTRT=…`,
  `LBORRES/LBSTRESC` — split on `/ = , ; : ( )`, not just whitespace.
- **Ranges.** `SVUPDE1 to SVUPDE12`, `SUDRINK1 to SUDRINK4 in SUPPSU` — expand them
  (also in the `<QNAM> in SUPPxx` scan).
- **Value-level rows** (`--ORRES` per `--TESTCD`): take the parent variable's
  own-domain pages that also mention the where-clause value
  (`LBORRES when LBTESTCD=HCG`), never every page of the parent.
- After changing the extractor, **re-run a previously verified study** and spot-check
  every page that changed directly in the PDF.

## The authority rule for page *values*
When the spec says a page and the aCRF says another, **the aCRF wins** — change
the page number, don't touch the aCRF. Example seen: spec said `AELNKID` on pages
`80 81` but the aCRF box color on p82 belonged to `HO`, so `AELNKID`'s real page
set excluded 82. Conversely, do not use a stale `[r_pages]` spec value as a
fallback except for a variable that is genuinely blank-page in the aCRF.

## SUPP-- pages
Derive SUPP qualifier pages the same way — the qualifier is annotated on the
parent form page. Update `SUPPxx` variable pages alongside the parent domain.

**Prefer the full-text `<QNAM> in SUPP<dom>` scan for supplemental qualifiers.**
The color extractor's RELREC / multi-line-span filters frequently **drop** a
qualifier that *is* annotated (e.g. `AECMNO in SUPPAE` sitting next to a `RELREC
when …` line on the same page). `scan_supp_qnam_pages(doc, offset)` scans the page
text for the `<QNAM> in SUPP<dom>` pattern and returns `{(SUPPxx, QNAM): pages}` —
typically recovering far more qualifier pages than the color extractor. Use it as the **primary** source for SUPP QVAL value-level pages, the color
extractor as fallback; only truly-absent qualifiers (ATC codes, device IMEI/serial,
vendor-derived) stay flagged (usually `Source = Vendor`, no page).

## Surgical corrections
When re-deriving pages for an already-good define, prefer a **per-(variable,page)
foreign-only removal** (remove only pages where the box color proves the
annotation belongs to another domain) over a full rebuild. A full rebuild can
mangle non-prefixed DM variables and is riskier than a targeted trim. Print a
before/after list of every changed page, but only **orange-highlight** the cells
that need review (a spec-vs-aCRF conflict) — not every routine page fill (see the
highlight rule in SKILL.md).

## Disagreement & conflict handling (Programmer Review)

Do not silently override the spec — preserve the original and flag it.

- **Spec page ≠ aCRF page** → **aCRF wins**: put the aCRF page in `Pages`, and
  write the original spec page + reason into the **`Programmer Review`** column
  (e.g. `Pages: spec=99, aCRF=85 → used 85`). The programmer reconciles.
- **aCRF annotates a variable whose spec origin is not CRF** (Derived/Assigned/
  eDT/Protocol) — a genuine annotation means it is collected, so this is a
  conflict. Do **not** change the origin; record it in `Programmer Review`
  (variable, aCRF page, spec origin) for the programmer to confirm.
  - *Exception (fill-empty):* if the spec origin is **blank** and the aCRF
    annotates it, set `Collected`/CRF and note it for a light confirm.
- Same rule for any override where data/aCRF disagrees with the spec: keep the
  accepted value, record the original + reason, let the programmer decide.

## Name-scan fallback (before flagging a missing page)

The color→domain filter is precise but can wrongly exclude a genuine annotation
(the box color didn't map to the variable's domain, or the header wasn't found on
that page). So when the filtered extractor returns **no page for a
Collected/Investigator variable**, do a **looser name-match scan** of the aCRF —
search the annotation-layer spans for the exact variable/QNAM token (still skip
NOTE boxes), and take the page(s) where it appears (`name_scan(doc, offset)` in
`acrf_annotations.py`). This typically recovers a handful of variables whose box
color did not map to their domain; note the page in `Programmer Review` for a light
confirm (the domain is not verified).
Only when the name-scan *also* finds nothing do you flag it — a truly-unannotated
Collected variable is usually vendor-assay data (IS/LC/PC results, units, ids,
EDC dates) that should be `Source = Vendor` (no page), which is the programmer's call.
