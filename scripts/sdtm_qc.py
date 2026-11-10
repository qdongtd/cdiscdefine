# -*- coding: utf-8 -*-
"""Systematic pre-check of an SDTM Define-XML 2.1 spec workbook against the
   spec-checkable subset of the Pinnacle 21 rule catalog.

   IMPORTANT: this is a PRE-FILTER to cut P21 round-trips, NOT authoritative.
   ~220 of the ~279 catalog rules need the *generated* define.xml, the CT
   database, MedDRA, the standard-label DB, or the data, and cannot be checked
   from the workbook. Always still generate define.xml and run P21.

   Lessons baked in (things an earlier hand-rolled checker got wrong):
     - Check the ValueLevel tab too, not just Variables (eDT, Derived-needs-Method).
     - DD0074: variable-level Origin must match every value-level Origin, else it
       should be blank (origin defined at value level only).
     - Length: flag NON-INTEGER length (page numbers leaked in), not just empty.
     - Codelist "missing" must also look in the Dictionaries tab (MedDRA/WHODD/LOINC).
     - SDTM Role set is Identifier/Topic/*Qualifier/Timing (NOT the ADaM roles).

     - Origin and Source must agree (Assigned/Derived/Protocol/Predecessor -> Sponsor).
     - Method/Comment text must be reviewer-readable: no librefs, SAS calls, raw
       names, spec-layout references, template placeholders or unbalanced quotes.
     - After any aCRF update, re-check EVERY page link: an update can replace an
       annotation (e.g. --DTC -> an EDC date in SUPP--) and orphan the old page.

   Usage:  python self_qc.py <define_spec.xlsx> [--acrf acrf.pdf] [--package DIR]
             --acrf     check every Pages cell against the aCRF annotations
                        (physical PDF page numbers)
             --package  check the submission folder: linked documents present,
                        no unaccepted tracked changes in reviewer's guides
"""
import sys, re, os, zipfile, warnings
import openpyxl
from collections import Counter, defaultdict
warnings.simplefilter('ignore')

ARGS = sys.argv[1:]
def opt(name):
    if name in ARGS:
        i = ARGS.index(name)
        v = ARGS[i + 1] if i + 1 < len(ARGS) else None
        del ARGS[i:i + 2]
        return v
    return None
ACRF, PKG = opt('--acrf'), opt('--package')
F = ARGS[0] if ARGS else 'define.xlsx'
wb = openpyxl.load_workbook(F)
def hdr(ws): return {n: i for i, n in enumerate([c.value for c in ws[1]])}
V, D, VL, CL = wb['Variables'], wb['Datasets'], wb['ValueLevel'], wb['Codelists']
M, C = wb['Methods'], wb['Comments']
ci, di, li, cli, mi, coi = hdr(V), hdr(D), hdr(VL), hdr(CL), hdr(M), hdr(C)
DICT = hdr(wb['Dictionaries']) if 'Dictionaries' in wb.sheetnames else None
vr = [r for r in V.iter_rows(min_row=2, values_only=True) if r[ci['Variable']]]
dr = [r for r in D.iter_rows(min_row=2, values_only=True) if r[di['Dataset']]]
vlr = [r for r in VL.iter_rows(min_row=2, values_only=True) if r[li['Dataset']]]
clr = [r for r in CL.iter_rows(min_row=2, values_only=True) if r[cli['ID']]]
DATE = {'datetime','date','time','partialDate','partialTime','partialDatetime','incompleteDatetime','durationDatetime','intervalDatetime'}
NUM = {'integer','float','text'}
res, ex = {}, defaultdict(list)
def chk(rule, viols):
    res[rule] = len(viols); ex[rule] = list(viols)[:4]
def is_int(x):
    try: int(str(x).strip()); return True
    except Exception: return False

# ---------- Datasets ----------
VALIDCLASS = {'SPECIAL PURPOSE','FINDINGS','FINDINGS ABOUT','EVENTS','INTERVENTIONS','TRIAL DESIGN','RELATIONSHIP'}
# Trial-design domains: their --DY/VISITDY etc. are Protocol-planned, NOT Derived -> exempt from DD0105.
TRIALDESIGN = {'TA','TE','TI','TV','TS','TD','TM'}
chk('DD0055 Class valid', ['%s:%s'%(r[di['Dataset']],r[di['Class']]) for r in dr if not str(r[di['Dataset']]).startswith('SUPP') and r[di['Class']] and r[di['Class']] not in VALIDCLASS])
chk('DD0054 Class required', [r[di['Dataset']] for r in dr if r[di['Class']] in (None,'')])
chk('MD0010 Structure required', [r[di['Dataset']] for r in dr if r[di['Structure']] in (None,'')])
chk('DD0040 Key Variables required', [r[di['Dataset']] for r in dr if r[di['Key Variables']] in (None,'')])
if 'Standard' in di:
    chk('Standard column filled', [r[di['Dataset']] for r in dr if r[di['Standard']] in (None,'')])
chk('DD0058 Dataset Label required', [r[di['Dataset']] for r in dr if r[di['Label']] in (None,'')])
if 'Repeating' in di:
    chk('MD0017 Repeating Y/N', [r[di['Dataset']] for r in dr if r[di['Repeating']] not in (None,'','Yes','No')])
chk('DD0135 Repeating=No if RefData=Yes', [r[di['Dataset']] for r in dr if 'Reference Data' in di and r[di['Reference Data']]=='Yes' and r[di['Repeating']]=='Yes'])

# ---------- Variables ----------
VALIDORIG = {'Collected','Derived','Assigned','Protocol','Predecessor'}
def V_(name): return ci[name]
def dtc(v): return v.endswith('DTC') or v.endswith('DUR')
chk('OD0075 DataType valid', ['%s.%s:%s'%(r[V_('Dataset')],r[V_('Variable')],r[V_('Data Type')]) for r in vr if r[V_('Data Type')] and r[V_('Data Type')] not in (DATE|NUM)])
chk('MD0034 DataType required', ['%s.%s'%(r[V_('Dataset')],r[V_('Variable')]) for r in vr if r[V_('Data Type')] in (None,'')])
chk('DD0149 --DTC/DUR date type', ['%s.%s'%(r[V_('Dataset')],r[V_('Variable')]) for r in vr if dtc(str(r[V_('Variable')])) and r[V_('Data Type')] not in DATE])
chk('OD0070 Length required(int/flt/txt)', ['%s.%s'%(r[V_('Dataset')],r[V_('Variable')]) for r in vr if r[V_('Data Type')] in NUM and r[V_('Length')] in (None,'')])
chk('OD0013 Length must be integer', ['%s.%s:%r'%(r[V_('Dataset')],r[V_('Variable')],r[V_('Length')]) for r in vr if r[V_('Length')] not in (None,'') and not is_int(r[V_('Length')])])
chk('DD0068 Length empty on non-num', ['%s.%s'%(r[V_('Dataset')],r[V_('Variable')]) for r in vr if r[V_('Data Type')] and r[V_('Data Type')] not in NUM and r[V_('Length')] not in (None,'')])
chk('OD0071 SigDig required(float)', ['%s.%s'%(r[V_('Dataset')],r[V_('Variable')]) for r in vr if r[V_('Data Type')]=='float' and r[V_('Significant Digits')] in (None,'')])
chk('DD0069 SigDig empty non-float', ['%s.%s'%(r[V_('Dataset')],r[V_('Variable')]) for r in vr if r[V_('Data Type')]!='float' and r[V_('Significant Digits')] not in (None,'')])
# DD0072: origin required, BUT optional at variable level when provided at value level
_vl_has_origin = {(r[li['Dataset']], r[li['Variable']]) for r in vlr if r[li['Origin']]}
chk('DD0072 Origin required', ['%s.%s'%(r[V_('Dataset')],r[V_('Variable')]) for r in vr
    if r[V_('Origin')] in (None,'') and r[V_('Variable')]!='QVAL'
    and (r[V_('Dataset')], r[V_('Variable')]) not in _vl_has_origin])
chk('DD0073 Origin valid(2.1)', ['%s.%s:%s'%(r[V_('Dataset')],r[V_('Variable')],r[V_('Origin')]) for r in vr if r[V_('Origin')] and r[V_('Origin')] not in VALIDORIG])
chk('DD0106 DOMAIN=Assigned', [r[V_('Dataset')] for r in vr if r[V_('Variable')]=='DOMAIN' and r[V_('Origin')]!='Assigned'])
chk('DD0107 RDOMAIN=Assigned', [r[V_('Dataset')] for r in vr if r[V_('Variable')]=='RDOMAIN' and r[V_('Origin')]!='Assigned'])
chk('DD0108 STUDYID=Protocol', [r[V_('Dataset')] for r in vr if r[V_('Variable')]=='STUDYID' and r[V_('Origin')]!='Protocol'])
chk('DD0105 studyday=Derived', ['%s.%s:%s'%(r[V_('Dataset')],r[V_('Variable')],r[V_('Origin')]) for r in vr if re.search(r'(DY|STDY|ENDY)$',str(r[V_('Variable')])) and r[V_('Dataset')] not in TRIALDESIGN and r[V_('Origin')] and r[V_('Origin')]!='Derived'])
chk('DD0042 Derived needs Method(var)', ['%s.%s'%(r[V_('Dataset')],r[V_('Variable')]) for r in vr if r[V_('Origin')]=='Derived' and r[V_('Method')] in (None,'')])
# Only Derived gets a Method: Collected/CRF/Assigned/Predecessor must NOT carry one
chk('Collected/CRF has spurious Method(var)', ['%s.%s'%(r[V_('Dataset')],r[V_('Variable')]) for r in vr if r[V_('Origin')] and r[V_('Origin')]!='Derived' and r[V_('Method')] not in (None,'')])
chk('DD0061 Predecessor needs value', ['%s.%s'%(r[V_('Dataset')],r[V_('Variable')]) for r in vr if r[V_('Origin')]=='Predecessor' and r[V_('Predecessor')] in (None,'')])
chk('DD0148 Predecessor syntax', ['%s.%s'%(r[V_('Dataset')],r[V_('Variable')]) for r in vr if r[V_('Origin')]=='Predecessor' and r[V_('Predecessor')] and not re.match(r'^[A-Za-z0-9_]+\.[A-Za-z0-9_]+', str(r[V_('Predecessor')]).strip())])
chk('DD0035 Pages req(Collected/Inv-Subj)', ['%s.%s'%(r[V_('Dataset')],r[V_('Variable')]) for r in vr if r[V_('Origin')]=='Collected' and r[V_('Source')] in ('Investigator','Subject') and r[V_('Pages')] in (None,'')])
chk('DD0111 Pages empty otherwise', ['%s.%s'%(r[V_('Dataset')],r[V_('Variable')]) for r in vr if r[V_('Pages')] not in (None,'') and not (r[V_('Origin')]=='Collected' and r[V_('Source')] in ('Investigator','Subject'))])
def uniq(var): return set(r[V_('Origin')] for r in vr if r[V_('Variable')]==var and r[V_('Origin')])
chk('DD0109 USUBJID consistent', [str(sorted(uniq('USUBJID')))] if len(uniq('USUBJID'))>1 else [])
chk('DD0110 VISITNUM consistent(Notice)', [str(sorted(uniq('VISITNUM')))] if len(uniq('VISITNUM'))>1 else [])
chk('DD0134 HasNoData needs Comment', ['%s.%s'%(r[V_('Dataset')],r[V_('Variable')]) for r in vr if str(r[V_('Has No Data')]).strip()=='Yes' and r[V_('Comment')] in (None,'')])

# ---------- ValueLevel (the tab an earlier checker forgot) ----------
chk('DD0073 Origin valid(VL)', ['%s.%s'%(r[li['Dataset']],r[li['Variable']]) for r in vlr if r[li['Origin']] and r[li['Origin']] not in VALIDORIG])
chk('DD0042 Derived needs Method(VL)', ['%s.%s'%(r[li['Dataset']],r[li['Variable']]) for r in vlr if r[li['Origin']]=='Derived' and r[li['Method']] in (None,'')])
chk('Collected/CRF has spurious Method(VL)', ['%s.%s'%(r[li['Dataset']],r[li['Variable']]) for r in vlr if r[li['Origin']] and r[li['Origin']]!='Derived' and r[li['Method']] not in (None,'')])
chk('DD0035 Pages req(VL Collected)', ['%s'%(r[li['Dataset']]) for r in vlr if r[li['Origin']]=='Collected' and r[li['Source']] in ('Investigator','Subject') and r[li['Pages']] in (None,'')])
chk('MD0064 WhereClause comparator', [r[li['Dataset']] for r in vlr if r[li['Where Clause']] and not re.search(r'\b(LT|LE|GT|GE|EQ|NE|IN|NOTIN)\b', str(r[li['Where Clause']]))])
# DD0074: variable-level origin must match ALL its value-level origins, else must be blank
vl_orig = defaultdict(set)
for r in vlr:
    if r[li['Origin']]: vl_orig[(r[li['Dataset']], r[li['Variable']])].add(r[li['Origin']])
d74 = []
for r in vr:
    k = (r[V_('Dataset')], r[V_('Variable')])
    if k in vl_orig and r[V_('Origin')] and {r[V_('Origin')]} != vl_orig[k]:
        d74.append('%s.%s var=%s vl=%s'%(k[0], k[1], r[V_('Origin')], sorted(vl_orig[k])))
chk('DD0074 var/value-level origin mismatch', d74)
# DD0123 value-level length <= variable length
vlen = {}
for r in vr:
    if is_int(r[V_('Length')]): vlen[(r[V_('Dataset')],r[V_('Variable')])] = int(r[V_('Length')])
d123 = []
for r in vlr:
    if is_int(r[li['Length']]):
        pv = vlen.get((r[li['Dataset']], r[li['Variable']]))
        if pv and int(r[li['Length']]) > pv: d123.append('%s.%s'%(r[li['Dataset']],r[li['Variable']]))
chk('DD0123 VL length<=var', d123)
supp_ds = {r[V_('Dataset')] for r in vr if str(r[V_('Dataset')]).startswith('SUPP')}
supp_vl = {r[li['Dataset']] for r in vlr if str(r[li['Dataset']]).startswith('SUPP')}
chk('DD0038 QVAL value-level per SUPP', sorted(supp_ds - supp_vl))

# ---------- Codelists ----------
seen = set(); byname = Counter()
for r in clr:
    if r[cli['ID']] not in seen:
        seen.add(r[cli['ID']])
        if r[cli['Name']]: byname[r[cli['Name']]] += 1
chk('DD0131 duplicate codelist Name', ['%s(x%d)'%(n,c) for n,c in byname.items() if c>1])
chk('DD0127 NCI code w/o Terminology', [r[cli['ID']] for r in clr if r[cli['NCI Codelist Code']] and not r[cli['Terminology']]])
# referenced codelist exists (codelists OR dictionaries)
clids = {r[cli['ID']] for r in clr}
dictids = {r[0] for r in wb['Dictionaries'].iter_rows(min_row=2, values_only=True) if r and r[0]} if DICT else set()
usedcl = {r[V_('Codelist')] for r in vr if r[V_('Codelist')]} | {r[li['Codelist']] for r in vlr if r[li['Codelist']]}
chk('DD0124 referenced codelist/dict missing', sorted(c for c in usedcl if c not in clids and c not in dictids))

# ---------- Methods / Comments / referential ----------
mids = {r[mi['ID']] for r in M.iter_rows(min_row=2, values_only=True) if r[mi['ID']]}
cids = {r[coi['ID']] for r in C.iter_rows(min_row=2, values_only=True) if r[coi['ID']]}
def dangling(idxm, idxc, rows):
    bad = []
    for r in rows:
        if r[idxm] and r[idxm] not in mids: bad.append('method %s'%r[idxm])
        if r[idxc] and r[idxc] not in cids: bad.append('comment %s'%r[idxc])
    return bad
chk('DD0016/DD0071 dangling Method/Comment', dangling(ci['Method'], ci['Comment'], vr) + dangling(li['Method'], li['Comment'], vlr))
# Method text left as an un-expanded pointer (should be the full METHODS-tab algorithm, domain-substituted)
POINTERS = {'--SEQ','--STDY','--ENDY','--DY','--LOBXFL','--BLFL','--SPID','--STRESN',
            '--ENRF','--ENRFTPT','--STRF','--STRFTPT','EPOCH','USUBJID','VISIT',
            'VISITNUM','VISITDY','TAETORD','RDOMAIN'}
def is_pointer(desc):
    d = str(desc or '').strip().upper()
    return d in POINTERS or bool(re.match(r'^--[A-Z]{2,8}$', d))
chk('Method text is un-expanded pointer', [r[mi['ID']] for r in M.iter_rows(min_row=2, values_only=True) if r[mi['ID']] and is_pointer(r[mi['Description']])])

# ---------- Method / Comment wording (reviewer-readable, no code) ----------
WORDING = [
    ('libref', re.compile(r'\b(sdtm|raw|adam|work)\.[A-Za-z_]+', re.I)),
    ('SAS function', re.compile(r'\b(catx|cats|catt|strip|put|input|substr|scan|compress|intck|intnx)\s*\(', re.I)),
    ('SAS format', re.compile(r'\bbest\d*\.(?!\d)', re.I)),
    ('spec-layout reference', re.compile(r'\b(tab|sheet)\s*["“]|\bkey variables column\b|\bsee (the )?(tab|sheet)\b', re.I)),
    ('placeholder', re.compile(r'"XX"|\bXX\b|\[if\b|\bTBD\b|\bTODO\b', re.I)),
    ('code syntax', re.compile(r'\bif\s+[A-Z][A-Z0-9_.]*\s*(=|\^=|~=|<>|\bEQ\b|\bNE\b|\bIN\b)[^.;]*\bthen\b|==|\bdo while\b')),
    ('bad character', re.compile(r'�')),
]
def wording(rows, idi, idd):
    bad = []
    for r in rows:
        t = str(r[idd] or '')
        hits = [n for n, p in WORDING if p.search(t)]
        if t.count('"') % 2:
            hits.append('unbalanced quotes')
        if hits:
            bad.append('%s (%s)' % (r[idi], ', '.join(hits)))
    return bad
mrows = [r for r in M.iter_rows(min_row=2, values_only=True) if r[mi['ID']]]
crows = [r for r in C.iter_rows(min_row=2, values_only=True) if r[coi['ID']]]
chk('Method wording not reviewer-readable', wording(mrows, mi['ID'], mi['Description']))
chk('Comment wording not reviewer-readable', wording(crows, coi['ID'], coi['Description']))

# ---------- Origin / Source agreement + Assigned-style text on Collected ----------
SPONSOR_ORIG = {'Assigned', 'Derived', 'Protocol', 'Predecessor'}
COLLECTED_ORIG = {'Collected', 'CRF', 'eDT'}
ctext = {r[coi['ID']]: str(r[coi['Description']] or '') for r in crows}
ASSIGNED_STYLE = re.compile(r'one[- ]?(on|to)[- ]?one', re.I)
COPY_STYLE = re.compile(r'^\s*set to\s+([A-Z]{2})?[A-Z]*ORRES[A-Z]*\s*\.?\s*$', re.I)
def src_rows(rows, ix, label):
    bad_src, bad_txt, copy = [], [], []
    for r in rows:
        k = '%s.%s%s' % (r[ix['Dataset']], r[ix['Variable']], (' [%s]' % r[ix['Where Clause']]) if 'Where Clause' in ix and r[ix['Where Clause']] else '')
        o = r[ix['Origin']]
        s = r[ix['Source']] if 'Source' in ix else None
        if s not in (None, ''):
            if o in SPONSOR_ORIG and s != 'Sponsor':
                bad_src.append('%s %s/%s' % (k, o, s))
            if o in COLLECTED_ORIG and s == 'Sponsor':
                bad_src.append('%s %s/%s' % (k, o, s))
        t = ctext.get(r[ix['Comment']], '') if r[ix['Comment']] else ''
        if o in COLLECTED_ORIG and t and ASSIGNED_STYLE.search(t):
            bad_txt.append('%s: %s' % (k, t[:60]))
        if t and COPY_STYLE.match(t) and o != 'Predecessor':
            copy.append('%s %s: %s' % (k, o, t[:40]))
    chk('Origin/Source mismatch(%s)' % label, bad_src)
    chk('Collected with Assigned-style Comment(%s)' % label, bad_txt)
    chk('Copy-of---ORRES not Predecessor(%s)' % label, copy)
src_rows(vr, ci, 'var')
src_rows(vlr, li, 'VL')

# ---------- TS (trial summary) ----------
ts_vl = {}
for r in vlr:
    if r[li['Dataset']] == 'TS' and r[li['Variable']] == 'TSVAL':
        m = re.search(r'TSPARMCD\s+EQ\s+["\']?(\w+)', str(r[li['Where Clause']] or ''))
        if m: ts_vl[m.group(1)] = r
if ts_vl:
    tsvars = {r[V_('Variable')] for r in vr if r[V_('Dataset')] == 'TS'}
    long_ = []
    for p, r in ts_vl.items():
        t = ctext.get(r[li['Comment']], '')
        if is_int(r[li['Length']]) and int(r[li['Length']]) > 200:
            long_.append('%s len %s' % (p, r[li['Length']]))
        if re.search(r'\bTSVAL1\b', t) and 'TSVAL1' not in tsvars:
            long_.append('%s uses TSVAL1 but TS has no TSVAL1 variable' % p)
    chk('TS TSVAL >200 needs TSVAL1..n', long_)
    # declared IG version (Standards tab, Datasets Standard column, or Define tab) vs TS SDTIGVER
    declared = set()
    if 'Standards' in wb.sheetnames:
        sh = hdr(wb['Standards'])
        for r in wb['Standards'].iter_rows(min_row=2, values_only=True):
            if r and 'Name' in sh and 'IG' in str(r[sh['Name']] or '').upper() and 'Version' in sh and r[sh['Version']]:
                declared.add(str(r[sh['Version']]).strip())
    if 'Standard' in di:
        for r in dr:
            m = re.search(r'IG\s*v?(\d+\.\d+)', str(r[di['Standard']] or ''), re.I)
            if m: declared.add(m.group(1))
    ig = ts_vl.get('SDTIGVER')
    ig_v = re.findall(r"(\d+\.\d+)", ctext.get(ig[li['Comment']], '') + ' ' + str(ig[li['Assigned Value']] or '')) if ig and 'Assigned Value' in li else []
    chk('TS SDTIGVER vs declared IG version', ['TS %s vs declared %s' % (ig_v, sorted(declared))] if ig_v and declared and not set(ig_v) & declared else [])

# ---------- optional: aCRF page links (physical PDF pages) ----------
if ACRF:
    import fitz
    ANNOT_FONTS = ('ArialMT', 'Arial-BoldMT', 'Arial')
    doc = fitz.open(ACRF)
    pagetext = {}
    for p in doc:
        parts = [(a.info.get('content') or '') for a in p.annots()]          # annotation objects
        for b in p.get_text('dict').get('blocks', []):                         # flattened annotation font layer
            for l in b.get('lines', []):
                for s in l.get('spans', []):
                    if s.get('font', '').startswith(ANNOT_FONTS):
                        parts.append(s['text'])
        pagetext[p.number + 1] = ' '.join(parts)
    def pg_bad(rows, ix, label):
        bad = []
        for r in rows:
            name = r[ix['Variable']]
            m = re.search(r'QNAM\s+EQ\s+["\']?(\w+)', str(r[ix['Where Clause']] or '')) if 'Where Clause' in ix else None
            if m: name = m.group(1)
            for pg in re.findall(r'\d+', str(r[ix['Pages']] or '')):
                pg = int(pg)
                if pg > doc.page_count or not re.search(r'\b%s\b' % re.escape(str(name)), pagetext.get(pg, '')):
                    bad.append('%s.%s p%d' % (r[ix['Dataset']], name, pg))
        return bad
    chk('aCRF page has no annotation for var', pg_bad(vr, ci, 'var') + pg_bad(vlr, li, 'VL'))

# ---------- optional: submission package ----------
if PKG:
    if 'Documents' in wb.sheetnames:
        dh = hdr(wb['Documents'])
        hrefs = [r[dh['Href']] for r in wb['Documents'].iter_rows(min_row=2, values_only=True) if r and 'Href' in dh and r[dh['Href']]]
        chk('Linked document missing from package', [h for h in hrefs if not os.path.exists(os.path.join(PKG, str(h)))])
    tracked = []
    for fn in os.listdir(PKG):
        if fn.lower().endswith('.docx') and not fn.startswith('~$'):
            try:
                x = zipfile.ZipFile(os.path.join(PKG, fn)).read('word/document.xml').decode('utf8', 'ignore')
                n = len(re.findall(r'<w:(?:ins|del)\b', x))
                if n: tracked.append('%s: %d tracked changes' % (fn, n))
            except Exception:
                pass
    chk('Reviewer guide has unaccepted tracked changes', tracked)

# ---------- report ----------
viol = sorted(k for k in res if res[k] > 0)
clean = sorted(k for k in res if res[k] == 0)
print('=== self-QC pre-check: %s ===' % F)
print('spec-checkable rules run: %d   |   clean: %d   |   with findings: %d' % (len(res), len(clean), len(viol)))
print('(PRE-FILTER ONLY - still generate define.xml and run P21; ~220 catalog rules need the generated XML / CT / MedDRA / data)')
print()
if viol:
    print('FINDINGS:')
    for k in viol: print('  X %-40s %4d  e.g. %s' % (k, res[k], ex[k]))
else:
    print('No spec-checkable findings.')
print()
print('CLEAN (%d): %s' % (len(clean), ', '.join(k.split()[0] for k in clean)))
