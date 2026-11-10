# Pinnacle 21 Define-XML 2.1 rule catalog — coverage map

A distilled index of the **279-rule** P21 Define-XML 2.1 validation catalog, used
to keep [`scripts/self_qc.py`](../scripts/self_qc.py) honest about what it does and
does not pre-check. This lists each rule's **ID and short title only** — not P21's
full rule text — as a coverage map, not a copy of the catalog. The authoritative
source is the P21 rule export (`define-xml 2.1.csv`); keep that **locally**, not in
this repo, and regenerate this file when the P21 / CT version bumps.

## Why most rules are not pre-checkable
`self_qc.py` reads the **Excel spec workbook only**. It cannot see the *generated*
define.xml (XML structure, OIDs, namespaces, hrefs), the CDISC CT / MedDRA
databases, the standard label DB, or the submitted `.xpt` data — so the ~240 rules
that need those stay P21-only. **Always still generate define.xml and run the P21
Define validation.** This pre-check just cuts round-trips on the spec-level rules.

## Coverage summary

| | Count |
|---|---|
| Total catalog rules | 279 |
| Implemented in `self_qc.py` | **38** |
| Not implemented | 241 |
| — of those, **candidate** spec-checkable (review) | 27 |
| — need generated XML / CT / MedDRA / data (P21-only) | 214 |

## Implemented (pre-checked by self_qc.py)

| ID | Rule (short title) |
|---|---|
| DD0016 | Referenced Method is missing |
| DD0035 | Missing Pages value |
| DD0038 | Missing Value Level metadata for QVAL in Dataset <dataset> |
| DD0040 | Missing Key Variables value |
| DD0042 | Missing Method reference |
| DD0054 | Missing Class value |
| DD0055 | Invalid Class value |
| DD0058 | Missing Label value |
| DD0061 | Missing Predecessor value |
| DD0068 | Invalid use of Length |
| DD0069 | Invalid use of Significant Digits |
| DD0071 | Referenced Comment is missing |
| DD0072 | Missing Origin |
| DD0073 | Invalid Origin Type value |
| DD0074 | Variable/Value Level Origin Type mismatch |
| DD0105 | Origin for Study Day variable <define variable> is not set to Derived |
| DD0106 | Origin for DOMAIN variable is not set to Assigned |
| DD0107 | Origin for RDOMAIN variable is not set to Assigned |
| DD0108 | Origin for STUDYID variable is not set to Protocol |
| DD0109 | Inconsistent Origin for USUBJID |
| DD0110 | Inconsistent Origin for VISITNUM |
| DD0111 | Invalid use of Pages |
| DD0123 | Variable/Value Level Length mismatch |
| DD0124 | Expected Codelist is missing |
| DD0127 | Invalid use of NCI Code for Codelist |
| DD0131 | Duplicate Codelist Name |
| DD0134 | Missing required def:CommentOID for Variable |
| DD0135 | Repeating/IsReferenceData mismatch |
| DD0148 | Invalid syntax for Predecessor |
| DD0149 | Invalid Data Type for --DTC/--DUR variable |
| MD0010 | Missing Structure value |
| MD0017 | Invalid Repeating value |
| MD0034 | Missing Data Type value |
| MD0064 | Missing Comparator value |
| OD0013 | Invalid integer value for <attribute> |
| OD0070 | Missing Length value |
| OD0071 | Missing Significant Digits value |
| OD0075 | Invalid Data Type value for variable |

## Candidate spec-checkable gaps — NOT yet in self_qc.py

Shortlist of rules whose title is an unambiguous **workbook-metadata** check
(no XML/CT/MedDRA/data dependency). **Confirm each before coding** — a title is a
filter, not a verdict. Note some are **sibling IDs of rules already implemented**
under a `DD####` number (P21 emits the same check at both a metadata `MD####` and
a define `DD####` tier): e.g. `MD0008`≈`DD0054` (Missing Class), `MD0009`≈`DD0055`
(Invalid Class), `MD0014`≈`DD0040` (Key Variables), `MD0016`≈`MD0017` (Repeating) —
coding the `DD` side already covers these. The genuinely new, cheap wins are the
string checks: name **uppercase**/**≤8 char** (MD0001/2/28/30), **Label ≤40**
(MD0006/33), **Structure sentence-case** (MD0011), **duplicate/missing Order**
(MD0025/36/77, OD0042, DD0077), and **DataType/Length vs Assigned Value** (DD0145/146).

| ID | Rule (short title) |
|---|---|
| DD0057 | Missing Description value |
| DD0077 | Missing Order number |
| DD0104 | Invalid Type value for Method |
| DD0145 | Data Type/Assigned Value mismatch |
| DD0146 | Length/Assigned Value mismatch |
| MD0001 | Dataset name is not uppercase |
| MD0002 | Dataset name longer than 8 characters |
| MD0005 | Missing Description value |
| MD0006 | Label longer than 40 characters |
| MD0008 | Missing Class value |
| MD0009 | Invalid Class value |
| MD0011 | Structure not in sentence case |
| MD0012 | Missing Purpose value |
| MD0014 | Missing Key Variables value |
| MD0016 | Missing Repeating value |
| MD0025 | Duplicate Order number |
| MD0026 | Variable is in wrong order within Dataset |
| MD0028 | Variable name is not uppercase |
| MD0030 | Variable name longer than 8 characters |
| MD0031 | Missing Label value |
| MD0033 | Label longer than 40 characters |
| MD0036 | Duplicate Order number |
| MD0041 | Codelist reference is not uppercase |
| MD0070 | Missing Description value and Document reference |
| MD0073 | Missing Description value and Document reference |
| MD0123 | Missing Description value |
| OD0042 | Duplicate Order number |

## Remaining rules (need generated define.xml / CT / MedDRA / data — P21-only)

<details><summary>214 rules — expand</summary>

| ID | Rule (short title) |
|---|---|
| DD0001 | XML schema validation issue within Define.xml |
| DD0002 | Missing or invalid <namespace> namespace reference |
| DD0003 | Missing required <attribute> value for <object> |
| DD0004 | Attribute <attribute> is not allowed to appear in element <element> |
| DD0005 | Attribute in wrong position within Define.xml |
| DD0006 | Missing required <element> value |
| DD0007 | Invalid content was found starting with element <element> |
| DD0008 | Element in wrong position within Define.xml |
| DD0009 | Element occurs more than once |
| DD0010 | ODM attribute <attribute> should not be included in Define.xml |
| DD0011 | ODM element <element> should not be included in Define.xml |
| DD0012 | Duplicate Document ID |
| DD0013 | Duplicate Method OID |
| DD0014 | Duplicate ValueListDef OID |
| DD0015 | Referenced Document is missing |
| DD0017 | Referenced Value Level metadata is missing |
| DD0018 | def:ArchiveLocationID/def:leaf mismatch |
| DD0019 | Invalid FileType value |
| DD0020 | Invalid def:DefineVersion value |
| DD0020A | Invalid def:DefineVersion value |
| DD0021 | Invalid Standard Name value <value> |
| DD0022 | Invalid Standard Version value <value> for <standard> |
| DD0023 | Element in wrong order within Define.xml |
| DD0024 | Invalid Term in codelist <codelist> for variable <variable> |
| DD0025 | Invalid MedDRA Version <version> |
| DD0026 | Required element 'Alias' is missing |
| DD0027 | C-Code not found in CDISC Controlled Terminology |
| DD0028 | Term/NCI Code mismatch in Codelist <codelist> |
| DD0029 | Required attribute def:ExtendedValue is missing or empty |
| DD0030 | Invalid def:ExtendedValue value |
| DD0032 | Missing NCI Code for Term in Codelist <codelist> |
| DD0033 | Unknown NCI Code value for Codelist <codelist> |
| DD0034 | Unknown NCI Code value for Term in Codelist <codelist> |
| DD0037 | Missing or invalid page range |
| DD0041 | Duplicate KeySequence |
| DD0044 | Invalid SoftHard value |
| DD0045 | Missing Domain value |
| DD0046 | Invalid use of Domain attribute |
| DD0047 | Missing SASDatasetName value |
| DD0048 | Invalid SASDatasetName value |
| DD0049 | Name/Domain/SASDatasetName mismatch |
| DD0050 | Domain/SASDatasetName mismatch for split dataset |
| DD0051 | Duplicate SASDatasetName |
| DD0052 | SASDatasetName/def:ArchiveLocationID mismatch |
| DD0053 | Invalid Purpose value |
| DD0056 | Missing def:ArchiveLocationID value |
| DD0062 | Duplicate xml:lang |
| DD0063 | Missing Alias |
| DD0064 | Invalid Alias Context value |
| DD0067 | Variable <variable> is not referenced |
| DD0070 | Missing SASFieldName value |
| DD0075 | CRF Origin is not referencing AnnotatedCRF |
| DD0076 | Missing Rank value |
| DD0078 | Document <document> is not referenced |
| DD0079 | Comment <comment> is not referenced |
| DD0080 | Method <method> is not referenced |
| DD0081 | Value Level metadata <value level> is not referenced |
| DD0082 | Codelist <codelist> is not referenced |
| DD0083 | Duplicate Comment OID |
| DD0084 | Referenced File is missing |
| DD0085 | Missing Define XSL |
| DD0086 | Invalid length of attributes in Define.xml |
| DD0087 | ARM metadata included in non ADaM define.xml |
| DD0088 | Duplicate arm:ResultDisplay OID |
| DD0089 | Duplicate arm:ResultDisplay Names |
| DD0090 | Missing Required Description Element |
| DD0091 | Duplicate arm:AnalysisResult OID |
| DD0092 | Missing required ParameterOID |
| DD0093 | Improper use of ParameterOID |
| DD0094 | Missing required def:CommentOID |
| DD0095 | Invalid ItemGroupOID reference |
| DD0096 | Missing required arm:AnalysisVariable element |
| DD0097 | Invalid ItemOID reference |
| DD0098 | Missing Context value |
| DD0099 | Extended value for AnalysisReason |
| DD0100 | Extended value for AnalysisPurpose |
| DD0102 | Invalid Annotated CRF document name |
| DD0103 | Missing referenced CRF document |
| DD0112 | Missing MedDRA definition |
| DD0113 | Missing MedDRA reference on <define variable> variable |
| DD0114 | Invalid usage of split datasets for non-general-observation-class datasets |
| DD0115 | Split and unsplit datasets are both listed |
| DD0116 | FATESTCD/FATEST mismatch in Codelist <codelist> |
| DD0117 | Incomplete Documentation for Analysis Results Metadata |
| DD0118 | NCI Code <define codelist code> for Codelist <define codelist> on Variable <define variable> does not match NCI Code <standard codelist code> for Standard Codelist <standard codelist> for Variable <standard variable> |
| DD0119 | Missing PublishingSet value |
| DD0120 | Invalid use of IsNonStandard for Dataset |
| DD0121 | Missing IsNonStandard attribute for Dataset |
| DD0122 | Referenced CDISC Standard is missing |
| DD0125 | Referenced CT Standard is missing |
| DD0126 | Invalid use of IsNonStandard for Codelist |
| DD0128 | Missing IsNonStandard attribute for Codelist |
| DD0129 | Missing required Alias |
| DD0130 | Invalid def:IsNonStandard value |
| DD0132 | Invalid use of def:ExtendedValue attribute |
| DD0133 | Missing required def:CommentOID for Dataset |
| DD0136 | Invalid dataset label value |
| DD0137 | Invalid variable label value |
| DD0138 | Referenced Standard is missing |
| DD0139 | Standard <standard> is not referenced |
| DD0140 | Invalid use of Subclass for Class |
| DD0141 | Invalid Status value |
| DD0147 | Assigned value not found in codelist |
| DD0150 | Missing expected Standard |
| DD0151 | Check Value not found in Codelist |
| DD0152 | Coded and Decoded values do not have the same C-Code |
| MD0003 | Analysis Dataset name doesn't start with 'AD' |
| MD0004 | Invalid characters in Dataset name |
| MD0007 | Description not in title case |
| MD0013 | Invalid Purpose value |
| MD0015 | Invalid Key Variables value |
| MD0018 | Missing Reference Data value |
| MD0019 | Invalid Reference Data value |
| MD0020 | Description value doesn't match base standard |
| MD0021 | Class value doesn't match base standard |
| MD0022 | Repeating value doesn't match base standard |
| MD0023 | Reference Data value doesn't match base standard |
| MD0024 | Missing Required or Expected variables |
| MD0027 | Invalid integer value |
| MD0029 | Invalid characters in Variable name |
| MD0032 | Label not in title case |
| MD0035 | Invalid Data Type value for variable |
| MD0037 | Invalid use of Length |
| MD0039 | Invalid use of Significant Digits |
| MD0040 | Missing Display Format value |
| MD0043 | Missing Expected Codelist ID value |
| MD0044 | Missing Expected Codelist Name value |
| MD0045 | Missing Origin value |
| MD0046 | Invalid Origin value |
| MD0047 | Missing Method reference |
| MD0048 | Missing Predecessor value |
| MD0049 | Invalid Role value |
| MD0050 | Missing Value Level metadata for QVAL |
| MD0051 | Variable/Value Level Origin mismatch |
| MD0052 | Missing Mandatory value |
| MD0053 | Invalid Mandatory value |
| MD0054 | Invalid Variable in dataset |
| MD0055 | Label value doesn't match base standard |
| MD0056 | Data Type value not compatible with base standard |
| MD0057 | Length greater than maximum allowed by base standard |
| MD0058 | Core value not compatible with base standard |
| MD0059 | Missing Codelist reference expected by base standard |
| MD0060 | Decoded Variable name doesn't match base standard |
| MD0061 | Origin value doesn't match base standard |
| MD0062 | Predecessor value doesn't match base standard |
| MD0063 | Role value doesn't match base standard |
| MD0065 | Invalid Comparator value |
| MD0066 | Missing WhereClause Value |
| MD0067 | Missing Name value |
| MD0068 | Missing Type value |
| MD0069 | Invalid Type value |
| MD0071 | Missing Expression Context value |
| MD0072 | Method is not referenced |
| MD0074 | Comment is not referenced |
| MD0075 | Missing Title value |
| MD0076 | Missing Href value |
| MD0077 | Document is not referenced |
| MD0114 | Codelist not found |
| MD0115 | Assigned Value not found in Codelist |
| MD0116 | Check Value not found in Codelist |
| MD0117 | Coded and Decoded values do not have the same C-Code |
| MD0118 | Invalid Codelist Name |
| MD0119 | Invalid Codelist Name |
| MD0120 | Invalid Codelist ID |
| MD0121 | Invalid Codelist Description |
| MD0122 | Invalid Codelist Description |
| MD0124 | Invalid Term value |
| MD0125 | Invalid Term value |
| MD0126 | Invalid Synonyms value |
| MD0127 | Invalid Synonyms value |
| MD0128 | Invalid Term value |
| MD0129 | Invalid Term value |
| MD0130 | Invalid Term value |
| MD0131 | Invalid Term value |
| MD0132 | Invalid Term value |
| MD0133 | Invalid Term value |
| MD0134 | Invalid Term value |
| MD0135 | Invalid Term value |
| MD0136 | Invalid Term value |
| MD0137 | Invalid Term value |
| OD0001 | XML is not well-formed |
| OD0002 | Missing or invalid Namespace attribute |
| OD0003 | Required attribute is missing or empty |
| OD0006 | Required element is missing or empty |
| OD0009 | Element occurs more than once |
| OD0010 | Missing XML declaration |
| OD0011 | Invalid XML encoding |
| OD0012 | Invalid root element |
| OD0014 | Invalid float value |
| OD0015 | Invalid date value |
| OD0016 | Invalid time value |
| OD0017 | Invalid datetime value for <attribute> |
| OD0018 | Invalid sasName value |
| OD0019 | Invalid SASFormatName value |
| OD0020 | Invalid fileName value |
| OD0021 | Invalid Language value |
| OD0022 | Duplicate Study OID |
| OD0027 | Duplicate MetaDataVersion OID |
| OD0030 | Duplicate ItemGroupDef OID |
| OD0031 | Duplicate ItemDef OID |
| OD0032 | Duplicate Codelist OID |
| OD0041 | Duplicate ItemRef ItemOID |
| OD0046 | Referenced ItemDef is missing |
| OD0048 | Referenced Codelist is missing |
| OD0072 | Invalid Repeating value |
| OD0073 | Invalid Reference Data value |
| OD0074 | Invalid Mandatory value |
| OD0076 | Invalid Data Type value for codelist |
| OD0077 | Codelist/Term Data Type mismatch |
| OD0078 | Invalid SASFormatName value for text Codelist |
| OD0079 | Duplicate Term in Codelist <codelist> |
| OD0080 | Codelist/Variable Data Type mismatch |
| OD0081 | Codelist <codelist> is empty |
| OD0082 | Codelist <codelist> is missing Decoded Values |

</details>
