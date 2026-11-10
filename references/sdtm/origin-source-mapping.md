# Origin / Source — version-aware

The valid **Origin** set depends on the **Define-XML version the define spec
declares** (Input 1 carries it — verify it). Use the matching set.

## Official definitions — CDISC Define-XML CT (the authority; quote these)

**Origin Type** — codelist **C170449**, submission value `ORIGINT`, **not
extensible** (CDISC Define-XML Controlled Terminology, NCI EVS; checked against the
2026-09-25 release):

| Code | Term | CDISC Definition (verbatim) | NCIt preferred term |
|---|---|---|---|
| C170548 | Collected | A value that is actually observed and recorded by a person or obtained by an instrument. | Collected Value |
| C170549 | Derived | A value that is calculated by an algorithm or reproducible rule, and which is dependent upon other data values. | Derived Value |
| C170547 | Assigned | A value that is derived through designation, such as values from a look up table or a label on a CRF. | Assigned Value |
| C170551 | Protocol | A value that is included as part of the study protocol. | Protocol Value |
| C170550 | Predecessor | A value that is copied from another variable. | Copied Value |
| C126101 | Not Available | A value that is not discoverable or accessible. | Not Available |
| C17649 | Other | Different than the one(s) previously specified or mentioned. (NCI) | Other |

**Origin Source** — codelist **C170450**: `Investigator`, `Sponsor`, `Subject`,
`Vendor` (see the CT for the full definitions).

Read the definitions literally — do **not** narrow them from memory:
- **Predecessor = copied from another variable.** The CT does **not** restrict it to
  another dataset; a same-dataset exact copy also fits. (Many guides/ADaM examples
  show cross-dataset copies — that is a usage pattern, not the definition.)
- **Derived = an algorithm or reproducible rule** applied to other data. Any
  conversion, standardization, mapping or computation makes it Derived.
- **Assigned = by designation**: a lookup table, a CRF label/pre-printed text,
  a sponsor-set constant or code (e.g. `--TESTCD` from the CRF question label,
  `DOMAIN`, dictionary/registry codes).
- **Protocol = stated in the study protocol** (planned design: TS design values,
  TA/TE/TV/TI, planned timepoints).

## Complete origin sets

| Standard / version | Valid Origin values |
|---|---|
| SDTM **Define-XML 2.0** | `CRF`, `Derived`, `Assigned`, `Protocol`, `eDT`, `Predecessor` |
| SDTM **Define-XML 2.1** | `Collected`, `Derived`, `Assigned`, `Protocol`, `Predecessor` |
| *(ref)* SEND 2.1 | adds `Other`, `Not Available` |
| *(ref)* ADaM 2.1 | only `Derived`, `Assigned`, `Predecessor` |

`eDT` and `CRF` are **not valid in Define-XML 2.1** — using them in a 2.1 define
triggers DD0073.

## Mapping (TD-spec origin/marker → define)

| TD-spec origin / marker | 2.0 Origin | 2.1 Origin | 2.1 Source | Fill |
|---|---|---|---|---|
| CRF, `[r_orig] CRF`, `[r_page]` (site-entered) | `CRF` | `Collected` | Investigator | Pages |
| CRF — subject-reported (ePRO, diary, QS) | `CRF` | `Collected` | Subject | Pages |
| eDT (central lab / e-transfer) | `eDT` | `Collected` | Vendor | — (no page) |
| Derived | `Derived` | `Derived` | Sponsor | **Method** |
| Assigned (designation: constants, lookup, CRF label, `--TESTCD`, units) | `Assigned` | `Assigned` | Sponsor | **Comment** |
| Protocol / trial design (STUDYID, TS) | `Protocol` | `Protocol` | Sponsor | — |
| Exact copy of another variable (same or other dataset, no change) | `Predecessor` | `Predecessor` | *(none — Source not used)* | Predecessor col `DATASET.VARIABLE` (DD0061/DD0148) |

### Predecessor vs Derived (e.g. `--STRESC` vs `--ORRES`)
SDTMIG describes `--STRESC` as "copied or derived from `--ORRES` in a standard
format or standard units", so either origin can be right:
- **Every record is an unchanged copy** → `Predecessor` (Predecessor =
  `VS.VSORRES`) fits the CT definition; `Derived` with Method `VSSTRESC = VSORRES`
  is also acceptable and is the common industry choice.
- **Any record is converted, standardized or mapped** (unit conversion, re-format,
  `NEG`→`NEGATIVE`) → `Derived` only, with the full rule as the Method.
Follow the TD spec's origin; if it is silent, **ask the user** for the team
convention and apply it consistently across domains (`--STRESC`/`--STRESN`).

### Complete 2.1 Source set
`Collected` pairs with **Investigator** (site/CRF), **Subject** (ePRO/diary/QS), or
**Vendor** (central lab, ECG/PK vendor — the old `eDT`). Everything else pairs with
**Sponsor**.

### Version notes
- **2.0** — `CRF`/`eDT` are valid; there is **no separate Source**. Pages required
  when Origin = `CRF`.
- **2.1** — a **Source** attribute is added. Pages required when Origin =
  `Collected` and Source = `Investigator` or `Subject`.

### House conventions
- STUDYID = `Protocol`/Sponsor; DOMAIN = `Assigned`/Sponsor; USUBJID =
  `Derived`/Sponsor with a `MT.<DS>.USUBJID` method (follow the other domains).
- Lab / vendor-assay domains (LB, LC, MB, PC, IS results) = `Collected`/`Vendor`.

## Method vs Comment — where the text comes from

- **Derived → Method** (`MT.<DS>.<VAR>`); **Assigned → Comment** (`COM.<DS>.<VAR>`).
- Text priority (copy verbatim): `[r_prog]` › **Conversion Definition (col N)** ›
  **Programming Notes (col O)**. *(N before O.)*
- If **N is too general** or **N ≠ O** substantially → still fill from the priority
  pick, but record **both N and O** in `Programmer Review` for the programmer to
  confirm.
- No source text → leave blank and flag in `Programmer Review`. Never invent.
- `Has No Data = Yes` variables → link a shared `COM.NODATA` ("planned but no data
  were collected…"), not a per-variable comment.

## aCRF annotation vs spec origin (conflict handling)

Only **collected** variables are annotated on the aCRF. So a genuine annotation
(own-domain colored box, passes the NOTE/structured/RELREC/color filters) is
evidence the variable is collected.

- **Spec origin blank** + aCRF annotated → set `Collected`/CRF (fill-empty), note
  in Programmer Review for a light confirm.
- **Spec origin explicitly non-CRF** (Derived/Assigned/eDT/Protocol) + aCRF
  annotated → **conflict.** Do **not** override; record it in `Programmer Review`
  (variable, aCRF page, spec origin) and let the programmer decide. Origin is a
  data-provenance decision the programmer owns.
