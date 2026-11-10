# -*- coding: utf-8 -*-
"""Fill a P21-generated SDTM Define-XML 2.1 spec workbook from the TD (compound)
SDTM spec (+ optional VLM) and the annotated CRF, following SKILL.md.

Study-agnostic: every study-specific value (study ID, versions, domains, pages,
text) comes from the inputs. Nothing is invented - an unsourced cell is left
blank, highlighted orange, and explained in a `Programmer Review` column.

    python fill_define.py --skeleton p21_define.xlsx --spec td_spec_sdtm.xlsx \
        --acrf acrf.pdf --out define_filled.xlsx [--vlm td_spec_vlm.xlsx] \
        [--study STUDYID] [--sibling other_study_define.xlsx] \
        [--page-ref physical|printed] [--subject-domains QS] [--ct YYYY-MM-DD]
        [--domain-vl inherit|flag]

--skeleton is never modified (it is the backup); --out is written fresh each run.
Run self_qc.py on the output afterwards, then generate define.xml and run P21.
"""
import re, sys, shutil, argparse, warnings, os
from collections import defaultdict, Counter
import openpyxl
from openpyxl.styles import PatternFill
import fitz
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import acrf_annotations as A
from common import ORG, NOFILL, DATE, SDTM_CLASSES, norm, txt, hdr, add_pr_col, put_pr
warnings.simplefilter('ignore')

ap = argparse.ArgumentParser()
ap.add_argument('--skeleton', required=True); ap.add_argument('--spec', required=True)
ap.add_argument('--acrf', required=True); ap.add_argument('--out', required=True)
ap.add_argument('--vlm'); ap.add_argument('--study', help='default: Define tab StudyName')
ap.add_argument('--sibling', help="a finished define of a sibling study: DataType/Length parity for blank skeleton cells")
ap.add_argument('--page-ref', choices=['physical', 'printed'], default='physical',
                help='physical (default) = actual PDF page number, what define.xml page links open; '
                     'printed = the CRF footer page number (cover offset removed)')
ap.add_argument('--subject-domains', default='QS', help='domains whose --ORRES is subject-reported (Source=Subject); "" = none')
ap.add_argument('--ct', help='CT version (YYYY-MM-DD) when the user overrides the spec Study tab CTVersion')
ap.add_argument('--domain-vl', choices=['inherit', 'flag', 'acrf'], default='inherit',
                help='skeleton domain value-level rows (not SUPP QVAL / TS) without VLM detail: inherit the parent '
                     'variable spec row; flag = leave blank for the programmer; acrf = per test: a Collected/'
                     'Investigator parent gives Collected + that test\'s own aCRF pages where the test is annotated '
                     '(else Vendor for --vendor-domains --ORRES, else blank + flag); other parents are inherited')
ap.add_argument('--vendor-domains', default='',
                help='with --domain-vl acrf: domains whose unannotated --ORRES tests come from a vendor transfer '
                     '(the spec source says so, e.g. a central-lab file) -> Collected/Vendor, flagged to confirm')
ap.add_argument('--nodata-text', default='Planned for collection but no data were collected for this study.')
args = ap.parse_args()

VALIDCLASS = SDTM_CLASSES
SUBJ_DOMS = {d.strip() for d in args.subject_domains.split(',') if d.strip()}

def pages_str(x): return ' '.join(re.findall(r'\d+', txt(x)))

try:
    shutil.copy(args.skeleton, args.out)
except PermissionError:   # --out is open in Excel
    args.out = args.out.replace('.xlsx', '_updated.xlsx'); shutil.copy(args.skeleton, args.out)
    print('NOTE: output open in Excel -> writing', args.out)
wb = openpyxl.load_workbook(args.out)
spec = openpyxl.load_workbook(args.spec, read_only=True, data_only=True)
STUDY = args.study or next(txt(r[1]) for r in wb['Define'].iter_rows(values_only=True) if txt(r[0]) == 'StudyName')
SK = norm(STUDY)
log, review = Counter(), []

# ---------------- TD spec: Study / Content / METHODS / domain tabs ----------------
def sheet(wbk, name):
    return next((wbk[s] for s in wbk.sheetnames if norm(s) == norm(name)), None)

def study_col(h):
    """index of the column whose header (or cached formula value) is this study"""
    return next((i for i, x in enumerate(h) if norm(x) == SK), None)

study_attrs = {}
st = sheet(spec, 'Study')
if st is not None:
    rows = list(st.iter_rows(values_only=True))
    k = next((i for i, x in enumerate(rows[0]) if norm(x) == SK), None)
    if k is None:  # header may be a var-style name; find the column whose StudyName row matches
        sn = next((r for r in rows if txt(r[0]) == 'StudyName'), None)
        k = next((i for i, x in enumerate(sn or []) if norm(x) == SK), None)
    if k is not None:
        study_attrs = {txt(r[0]): txt(r[k]) for r in rows[1:] if r[0]}
CTVER = args.ct or study_attrs.get('CTVersion', '')
if args.ct and study_attrs.get('CTVersion') and args.ct != study_attrs['CTVersion']:
    review.append('CT: user override %s (spec Study tab says %s) - align TS TSVCDVER too' % (args.ct, study_attrs['CTVersion']))
IGVER = (study_attrs.get('StandardVersion', '').split('/')[-1]).strip()
if not CTVER or not IGVER:
    review.append('Study tab: CTVersion/StandardVersion not found for %s - Standards tab not auto-filled' % STUDY)

cont = list(sheet(spec, 'Content').iter_rows(values_only=True))
ch = hdr(cont[0]); cst = study_col(cont[0])
if cst is None: sys.exit('Content tab has no column for study %s' % STUDY)
content = {txt(r[0]): r for r in cont[1:] if r[0]}
in_scope = {d for d, r in content.items() if txt(r[cst])}   # 'x' or a marker such as [r_sort]

mt = sheet(spec, 'METHODS')
methods_tab = {}
if mt is not None:
    for r in list(mt.iter_rows(values_only=True))[1:]:
        if r[0]:  # rejoin '--' split across a source line wrap before any substitution
            methods_tab[txt(r[0])] = txt(r[1]).replace('-\n-', '--').replace('--\n', '--')

R_ORIG = re.compile(r'\[r_orig\]\s*([A-Za-z]+)')
R_PAGE = re.compile(r'\[r_pages?\]\s*([0-9][0-9 ,]*)')
R_PROG = re.compile(r'\[r_prog\]\s*(.*?)(?=\[r_|$)', re.S)
specvars, specsupp = {}, defaultdict(dict)
for tab in spec.sheetnames:
    if tab not in in_scope: continue
    rows = list(spec[tab].iter_rows(values_only=True))
    if not rows or norm(rows[0][0]) != 'variablename': continue
    h = hdr(rows[0]); sc = study_col(rows[0])
    if sc is None: review.append('%s tab: no column for %s' % (tab, STUDY)); continue
    ni = h.get('programmingnotes', h.get('sascodes'))
    for r in rows[1:]:
        v = txt(r[0])
        if not v: continue
        s = txt(r[sc])
        d = dict(var=v, study=s, origin=txt(r[h['origin']]), pages=pages_str(r[h['pages']]),
                 conv=txt(r[h['conversiondefinition']]), notes=txt(r[ni]) if ni is not None else '',
                 cat=txt(r[h['variablecategory']]), type=txt(r[h['type']]) if 'type' in h else '',
                 length=r[h['length']] if 'length' in h else None)
        m = R_ORIG.search(s); d['r_orig'] = m.group(1) if m else ''
        m = R_PAGE.search(s); d['r_page'] = pages_str(m.group(1)) if m else ''
        m = R_PROG.search(s); d['r_prog'] = m.group(1).strip() if m else ''
        tgt, key = (specsupp[tab], v) if d['cat'] == 'S' else (specvars, (tab, v))
        if key in tgt and tgt[key]['study'] and not s: continue   # duplicate rows: keep the study-marked one
        tgt[key] = d

# ---------------- aCRF ----------------
doc = fitz.open(args.acrf)
cover = A.detect_cover_offset(doc)          # front-matter pages before printed page 1
offset = cover if args.page_ref == 'printed' else 0
NPAGES = doc.page_count if args.page_ref == 'physical' else doc.page_count - cover
acrf_bd = A.collapse_by_domain(A.extract(args.acrf, offset))
acrf_supp = A.scan_supp_qnam_pages(doc, offset)
acrf_ns = A.name_scan(doc, offset)
VENDOR_DOMS = {d.strip() for d in args.vendor_domains.split(',') if d.strip()}
PAGETEXT = {}   # annotation-layer text per page (same numbering as acrf_bd), for value-level where values
for _p in range(doc.page_count):
    PAGETEXT[_p - offset + 1] = ' '.join(s['text'] for b in doc[_p].get_text('dict')['blocks'] for l in b.get('lines', [])
                                         for s in l['spans'] if s['font'].startswith(A.ANNOT_FONTS)).upper()

# The TD spec's Pages may be numbered either way. Detect its basis from variables that
# have both a spec page and an own-domain aCRF page, then shift spec pages to --page-ref.
def _spec_pages(d): return [int(x) for x in (d['r_page'] or d['pages']).split()] if d else []
votes = Counter()
for (ds, v), d in specvars.items():
    sp, ap_ = _spec_pages(d), acrf_bd.get((v, A.norm_dom(ds)))
    if sp and ap_:
        for delta in {0, cover, -cover}:
            if sorted({p + delta for p in sp}) == ap_: votes[delta] += 1
SPEC_SHIFT = votes.most_common(1)[0][0] if votes else 0
SPEC_MIXED = bool(votes) and votes.most_common(1)[0][1] < 0.8 * sum(votes.values())
# No vote at all although both sides have pages: the spec was paged against a different
# aCRF version/layout (forms moved, varying offsets). Its pages cannot be mapped - never
# copy them into the define; flag the variable instead.
_both = sum(1 for (ds, v), d in specvars.items() if _spec_pages(d) and acrf_bd.get((v, A.norm_dom(ds))))
SPEC_STALE = not votes and _both >= 5
if SPEC_STALE:
    review.append('aCRF: TD spec pages match no consistent offset of this aCRF (%d variables compared) - spec was paged '
                  'against another aCRF version; spec-only pages NOT used, flagged instead' % _both)
if SPEC_MIXED:   # e.g. some [r_page] overrides already PDF-numbered, others printed
    review.append('aCRF: TD spec page numbering is MIXED (votes %s) - spec pages NOT shifted; '
                  'spec-only fallback pages flagged as basis-uncertain' % dict(votes))
    SPEC_SHIFT = 0
review.append('aCRF: page-ref=%s (PDF has %d pages, %d cover page(s)); TD spec page shift %s (votes %s) - verify one clean variable'
              % (args.page_ref, doc.page_count, cover, 'none (mixed)' if SPEC_MIXED else '%+d' % SPEC_SHIFT, dict(votes)))

def spec_to_ref(sp): return ' '.join(str(int(x) + SPEC_SHIFT) for x in sp.split()) if sp else ''

def acrf_pages(var, dom, supp=None):
    if supp and acrf_supp.get((supp, var)): return acrf_supp[(supp, var)], 'supp-scan'
    p = acrf_bd.get((var, A.norm_dom(dom)))
    if p: return p, 'color'
    p = acrf_ns.get(var)
    return (p, 'name-scan') if p else ([], '')

# ---------------- text / origin rules ----------------
def stem(var, dom):
    if var in methods_tab: return var
    if var.startswith(dom) and '--' + var[len(dom):] in methods_tab: return '--' + var[len(dom):]
    return None

def text_for(d, dom, var):
    """[r_prog] > Conversion Definition (N) > Programming Notes (O); expand METHODS pointers;
    name/stem fallback. Returns (text, flag_msgs, needs_orange)."""
    d = d or {}
    src, t = next(((l, c) for l, c in (('r_prog', d.get('r_prog')), ('N', d.get('conv')), ('O', d.get('notes'))) if c), ('', ''))
    msgs, orange = [], False
    if t and t.strip() in methods_tab:
        k = t.strip(); t = methods_tab[k].replace('--', dom)
        vs = stem(var, dom)
        if k.startswith('--') and vs and vs != k:
            msgs.append('spec pointer %s does not match variable stem %s; O="%s"' % (k, vs, d.get('notes', '')[:150]))
            orange = True
    elif not t:
        k = stem(var, dom)
        if k: t = methods_tab[k].replace('--', dom)
    if src == 'N' and d.get('notes') and norm(d['notes']) != norm(d['conv']) and not orange \
            and not d['notes'].lower().startswith('raw.'):
        msgs.append('N/O differ - N used; O="%s"' % d['notes'][:150])
    return t, msgs, orange

def map_origin(o):
    return {'crf': 'Collected', 'collected': 'Collected', 'edt': 'Collected', 'derived': 'Derived',
            'assigned': 'Assigned', 'protocol': 'Protocol', 'predecessor': 'Predecessor'}.get(norm(o), '')
def map_source(o, dom, var):
    o = norm(o)
    if o == 'edt': return 'Vendor'
    if o in ('crf', 'collected'):
        return 'Subject' if dom in SUBJ_DOMS and var == dom + 'ORRES' else 'Investigator'
    return 'Sponsor' if o else ''

methods, comments = {}, {}
COMMON = {'COM.NODATA': args.nodata_text,
          'COM.SUPP.RDOMAIN': 'Domain abbreviation from where data originated.',
          'COM.SUPP.IDVAR': 'Name of the variables for the related records.',
          'COM.SUPP.IDVARVAL': 'Value of identifying variable described in IDVAR.',
          'COM.SUPP.QVLM': 'Assigned based on origin described in QVAL VLM section'}
SUPP_FIX = {'RDOMAIN': 'COM.SUPP.RDOMAIN', 'IDVAR': 'COM.SUPP.IDVAR', 'IDVARVAL': 'COM.SUPP.IDVARVAL',
            'QNAM': 'COM.SUPP.QVLM', 'QLABEL': 'COM.SUPP.QVLM', 'QORIG': 'COM.SUPP.QVLM'}


def fill_row(ws, ci, row, ds, var, d, dom, supp=None, idsuffix=None):
    """Origin/Source/Pages/Method/Comment for one Variables or ValueLevel row (fill-empty-only)."""
    pr = []
    cell = lambda n: ws.cell(row=row, column=ci[n] + 1)
    cur = txt(cell('Origin').value)
    spec_o = (d['r_orig'] or d['origin']) if d else ''
    so = map_origin(spec_o)
    if cur in ('eDT', 'CRF'):   # Define-XML 2.0 origins -> 2.1
        cell('Origin').value = 'Collected'
        if not txt(cell('Source').value): cell('Source').value = 'Vendor' if cur == 'eDT' else 'Investigator'
        log['%s->Collected' % cur] += 1; cur = 'Collected'
    if not cur and so:
        cell('Origin').value = cur = so; log['origin filled'] += 1
    elif cur and so and cur != so:
        pr.append('Origin: skeleton=%s, spec=%s -> kept skeleton (fill-empty-only), confirm' % (cur, spec_o))
        cell('Origin').fill = ORG; log['origin conflict'] += 1
    elif not cur:
        pr.append('Origin: no source in TD spec for this study'); cell('Origin').fill = ORG; log['origin unsourced'] += 1
    if cur and not txt(cell('Source').value):
        cell('Source').value = (map_source(spec_o, dom, var) if so == 'Collected' else 'Investigator') if cur == 'Collected' else 'Sponsor'
    src = txt(cell('Source').value)
    if cur == 'Collected' and src in ('Investigator', 'Subject') and not txt(cell('Pages').value):
        raw_sp = (d['r_page'] or d['pages']) if d else ''
        sp = spec_to_ref(raw_sp)
        ap_, how = acrf_pages(var, dom, supp)
        aps = ' '.join(map(str, ap_))
        if aps:
            cell('Pages').value = aps; log['pages from aCRF'] += 1
            if sp and sp != aps:
                pr.append('Pages: spec=%s%s, aCRF=%s -> used aCRF'
                          % (raw_sp, '' if SPEC_SHIFT == 0 else ' (=%s after %+d shift)' % (sp, SPEC_SHIFT), aps))
            if how == 'name-scan': pr.append('Pages: color filter found none; name-scan fallback p%s - confirm domain' % aps)
        elif sp and any(int(x) > NPAGES or int(x) < 1 for x in sp.split()):
            cell('Pages').fill = ORG; log['pages spec out of range'] += 1
            pr.append('Pages: not found on aCRF; spec page(s) %s outside the %d-page aCRF (stale) - not used' % (raw_sp, NPAGES))
        elif sp and SPEC_STALE:
            cell('Pages').fill = ORG; log['pages spec-only, stale (not used)'] += 1
            pr.append('Pages: not found on aCRF; spec page %s is from another aCRF version - not used; annotate on aCRF or '
                      'confirm origin' % raw_sp)
        elif sp:
            cell('Pages').value = sp; cell('Pages').fill = ORG; log['pages spec-only'] += 1
            pr.append('Pages: not found on aCRF (color+name scan); spec page %s%s kept - confirm%s'
                      % (raw_sp, '' if SPEC_SHIFT == 0 else ' (-> %s after %+d shift)' % (sp, SPEC_SHIFT),
                         '; spec page numbering is mixed (printed vs PDF) - verify the page' if SPEC_MIXED else ''))
        else:
            cell('Pages').fill = ORG; log['pages missing'] += 1
            pr.append('Pages: none in spec or aCRF - likely vendor data (Source=Vendor?) - programmer to decide')
    idb = '%s.%s' % (ds, idsuffix or var)
    if cur == 'Derived' and not txt(cell('Method').value):
        t, msgs, orange = text_for(d, dom, var)
        if t:
            methods['MT.' + idb] = t; cell('Method').value = 'MT.' + idb; log['method'] += 1
            pr += ['Method: ' + m for m in msgs]
            if orange: cell('Method').fill = ORG; log['method pointer mismatch'] += 1
        else:
            cell('Method').fill = ORG; log['method missing'] += 1
            pr.append('Method: no text in spec ([r_prog]/N/O empty, no METHODS match)')
    if cur == 'Assigned' and not txt(cell('Comment').value) and d:
        t, msgs, orange = text_for(d, dom, var)
        if t:
            comments['COM.' + idb] = t; cell('Comment').value = 'COM.' + idb; log['comment'] += 1
            pr += ['Comment: ' + m for m in msgs]
            if orange: cell('Comment').fill = ORG
    return pr

def nodata(ws, ci, rr):
    c = ws.cell(rr, ci['Comment'] + 1)
    if txt(ws.cell(rr, ci['Has No Data'] + 1).value) == 'Yes' and not txt(c.value):
        comments['COM.NODATA'] = COMMON['COM.NODATA']; c.value = 'COM.NODATA'; log['nodata comment'] += 1

# ---------------- Variables ----------------
V = wb['Variables']; vci = add_pr_col(V)
vrow = {}
for rr in range(2, V.max_row + 1):
    ds, var = txt(V.cell(rr, vci['Dataset'] + 1).value), txt(V.cell(rr, vci['Variable'] + 1).value)
    if not ds or not var: continue
    vrow[(ds, var)] = rr
    cell = lambda n: V.cell(rr, vci[n] + 1)
    if (ds.startswith('SUPP') and var in SUPP_FIX) or (ds == 'RELREC' and var in ('RDOMAIN', 'IDVAR', 'IDVARVAL')):
        if not txt(cell('Origin').value): cell('Origin').value = 'Assigned'
        if txt(cell('Origin').value) == 'Assigned':
            cid = SUPP_FIX[var]; comments[cid] = COMMON[cid]; cell('Comment').value = cid
            cell('Source').value = cell('Source').value or 'Sponsor'; log['supp template comment'] += 1
        else:
            put_pr(V, vci, rr, ['Origin %s != SUPP template Assigned' % cell('Origin').value]); cell('Origin').fill = ORG
        continue
    if ds.startswith('SUPP'):
        if var == 'QVAL': continue                    # origin lives at value level
        if var == 'QEVAL':                            # comment comes from source spec only
            if not txt(cell('Origin').value): cell('Origin').value = 'Assigned'
            cell('Source').value = cell('Source').value or 'Sponsor'; nodata(V, vci, rr); continue
        # SUPP template (SKILL.md): STUDYID = Protocol, USUBJID = Derived (METHODS-tab USUBJID)
        tmpl = {'STUDYID': 'Protocol', 'USUBJID': 'Derived'}.get(var, '')
        d = {'origin': tmpl, 'r_orig': '', 'r_prog': '', 'conv': var, 'notes': '', 'pages': '', 'r_page': ''}
        put_pr(V, vci, rr, fill_row(V, vci, rr, ds, var, d, ds)); continue
    d = specvars.get((ds, var)); pr = []
    if ds not in in_scope: pr.append('Dataset %s not marked for %s on Content tab' % (ds, STUDY))
    elif d is None: pr.append('Variable not found in TD %s tab' % ds)
    elif not d['study']: pr.append('Variable not marked for %s in TD spec (present in data)' % STUDY)
    pr += fill_row(V, vci, rr, ds, var, d, ds)
    nodata(V, vci, rr); put_pr(V, vci, rr, pr)

# DataType / Length backfill (MD0034 / OD0070) and date types (DD0149 / DD0068)
sib = {}
if args.sibling:
    sv = list(openpyxl.load_workbook(args.sibling, read_only=True)['Variables'].iter_rows(values_only=True))
    sh = {n: i for i, n in enumerate(sv[0])}
    sib = {(r[sh['Dataset']], r[sh['Variable']]): (r[sh['Data Type']], r[sh['Length']]) for r in sv[1:] if r[sh['Dataset']]}
SPECTYPE = {'text': 'text', 'char': 'text', 'integer': 'integer', 'num': None, 'float': 'float'}
def _spec_dtype(d):
    t = (d or {}).get('type', '').strip()
    if not t: return None
    dt_ = t if t in DATE else SPECTYPE.get(t.lower())
    if not dt_: return None
    L = d.get('length')
    return (dt_, None if dt_ in DATE else (str(int(L)) if isinstance(L, (int, float)) or str(L).strip().isdigit() else None))
struct = {}   # populated SUPP/RELREC structural vars in this define: var -> (type, max length)
for (ds, var), rr in vrow.items():
    if (ds.startswith('SUPP') or ds == 'RELREC') and txt(V.cell(rr, vci['Data Type'] + 1).value):
        t, L = V.cell(rr, vci['Data Type'] + 1).value, V.cell(rr, vci['Length'] + 1).value
        o = struct.get(var)
        if not o or (str(L).isdigit() and (not str(o[1]).isdigit() or int(L) > int(o[1]))): struct[var] = (t, L)
for (ds, var), rr in vrow.items():
    dt, ln = V.cell(rr, vci['Data Type'] + 1), V.cell(rr, vci['Length'] + 1)
    if not txt(dt.value):
        x = sib.get((ds, var)); how = 'sibling define'
        if (ds.startswith('SUPP') or ds == 'RELREC') and var in struct:
            x, how = struct[var], 'populated %s in another SUPP/RELREC dataset of this define (max length)' % var
        elif _spec_dtype(specvars.get((ds, var))):
            x, how = _spec_dtype(specvars.get((ds, var))), 'TD spec Type/Length'
        if x and x[0]:
            dt.value = x[0]; ln.value = None if x[0] in DATE else (str(x[1]) if x[1] is not None else None)
            put_pr(V, vci, rr, ['DataType/Length: blank in skeleton; %s/%s from %s' % (x[0], ln.value, how)]); log['dtype backfilled'] += 1
            continue
        if x and x[0]:
            dt.value = x[0]; ln.value = None if x[0] in DATE else (str(x[1]) if x[1] is not None else None)
            put_pr(V, vci, rr, ['DataType/Length: blank in skeleton; taken from sibling define (%s/%s)' % x]); log['dtype from sibling'] += 1
        else:
            dt.fill = ORG; put_pr(V, vci, rr, ['DataType: blank in skeleton, no source']); log['dtype unsourced'] += 1
    elif (var.endswith('DTC') or var.endswith('DUR')) and dt.value not in DATE:
        put_pr(V, vci, rr, ['DataType: skeleton=%s/%s -> partialDate, Length cleared (DD0149/DD0068) - confirm' % (dt.value, ln.value)])
        dt.value = 'partialDate'; ln.value = None; log['DD0149 date type'] += 1

# ---------------- ValueLevel ----------------
VL = wb['ValueLevel']; lci = add_pr_col(VL)
for rr in range(2, VL.max_row + 1):
    ds, var = txt(VL.cell(rr, lci['Dataset'] + 1).value), txt(VL.cell(rr, lci['Variable'] + 1).value)
    wc = txt(VL.cell(rr, lci['Where Clause'] + 1).value)
    if not ds: continue
    if ds.startswith('SUPP') and var == 'QVAL':
        m = re.match(r'QNAM EQ (\S+)', wc); q = m.group(1) if m else ''
        par = ds[4:]; d = specsupp.get(par, {}).get(q)
        pr = []
        if d is None: pr.append('QNAM %s not found as S-category row in TD %s tab' % (q, par))
        elif not d['study']: pr.append('QNAM %s not marked for %s in TD spec (present in data)' % (q, STUDY))
        pr += fill_row(VL, lci, rr, ds, q, d, par, supp=ds, idsuffix='QVAL.' + q)
        put_pr(VL, lci, rr, pr)
    elif args.domain_vl == 'flag' and ds != 'TS':   # no study VLM: leave for the programmer
        c = VL.cell(rr, lci['Origin'] + 1)
        if not txt(c.value):
            c.fill = ORG; log['domain VL left blank (flag)'] += 1
            put_pr(VL, lci, rr, ['Value level: no VLM detail for %s - Origin/Source/Pages/Method/Comment left for programmer' % STUDY])
        else:
            put_pr(VL, lci, rr, fill_row(VL, lci, rr, ds, var, specvars.get((ds, var)), ds,
                   idsuffix='%s.%s' % (var, re.sub(r'[^A-Za-z0-9]+', '_', wc)[:40].strip('_'))))
    elif args.domain_vl == 'acrf' and ds != 'TS' and not txt(VL.cell(rr, lci['Origin'] + 1).value):
        d = specvars.get((ds, var)); spec_o = (d['r_orig'] or d['origin']) if d else ''
        o, s_ = map_origin(spec_o), map_source(spec_o, ds, var)
        cell = lambda n: VL.cell(rr, lci[n] + 1)
        if o == 'Collected' and s_ in ('Investigator', 'Subject'):
            mval = re.match(r'\s*\S+\s+EQ\s+(.+?)(?:\s+and\s+.*)?$', wc, re.I)
            val = mval.group(1).strip(' "\'').upper() if mval else ''
            par = acrf_bd.get((var, A.norm_dom(ds))) or []
            hit = [p for p in par if val and re.search(r'(?<![A-Z0-9])%s(?![A-Z0-9])' % re.escape(val), PAGETEXT[p])]
            if hit:
                cell('Origin').value, cell('Source').value, cell('Pages').value = 'Collected', s_, ' '.join(map(str, hit))
                put_pr(VL, lci, rr, ['Collected: aCRF pages annotating %s for %s' % (var, val)]); log['domain VL Collected + pages'] += 1
            elif var.endswith('ORRES') and ds in VENDOR_DOMS:
                cell('Origin').value, cell('Source').value = 'Collected', 'Vendor'; cell('Origin').fill = ORG
                put_pr(VL, lci, rr, ['CONFIRM: %s not annotated on aCRF; %s results from vendor transfer per spec -> Collected/Vendor' % (val, ds)])
                log['domain VL Vendor (confirm)'] += 1
            else:
                cell('Origin').fill = ORG; log['domain VL not annotated (blank)'] += 1
                put_pr(VL, lci, rr, ['Parent Collected/%s but %s not annotated on aCRF - left blank for programmer' % (s_, val)])
        else:
            put_pr(VL, lci, rr, fill_row(VL, lci, rr, ds, var, d, ds,
                   idsuffix='%s.%s' % (var, re.sub(r'[^A-Za-z0-9]+', '_', wc)[:40].strip('_'))))
            log['domain VL inherited (%s)' % (o or 'none')] += 1
    else:  # domain value level with no VLM detail: inherit the parent variable's spec row
        pr = fill_row(VL, lci, rr, ds, var, specvars.get((ds, var)), ds,
                      idsuffix='%s.%s' % (var, re.sub(r'[^A-Za-z0-9]+', '_', wc)[:40].strip('_')))
        put_pr(VL, lci, rr, pr)

def where(wv, cm, wval):
    out = []
    for v, c, x in zip(*[[p.strip() for p in s.split(';')] for s in (wv, cm, wval)]):
        if c == 'EQ' and re.search(r'[\s,]', x): x = '"%s"' % x if "'" in x else "'%s'" % x
        out.append('%s %s %s' % (v, c, x))
    return ' AND '.join(out)

if args.vlm:
    vlm = openpyxl.load_workbook(args.vlm, read_only=True, data_only=True)
    order = VL.max_row - 1
    for tab in vlm.sheetnames:
        rows = list(vlm[tab].iter_rows(values_only=True))
        if tab not in in_scope or not rows or norm(rows[0][0]) != 'variablename': continue
        h = hdr(rows[0]); sc = study_col(rows[0])
        if sc is None:
            review.append('VLM %s: no column for %s - value level left for programmer (do not reuse another study)' % (tab, STUDY)); continue
        for r in rows[1:]:
            if not r[0]: continue
            if not txt(r[sc]): log['VLM rows not marked (skipped)'] += 1; continue
            var = txt(r[0]); vr = vrow.get((tab, var))
            if not vr: review.append('VLM %s.%s: variable not in define' % (tab, var)); continue
            order += 1; nr = VL.max_row + 1; pr = []
            put = lambda n, v: VL.cell(nr, lci[n] + 1, v)
            pv = lambda n: V.cell(vr, vci[n] + 1).value
            put('Order', str(order)); put('Dataset', tab); put('Variable', var)
            put('Where Clause', where(txt(r[h['wherevarlist']]), txt(r[h['comparator']]), txt(r[h['wherevalue']])))
            put('Data Type', pv('Data Type')); put('Mandatory', 'No')
            L, PL = r[h['length']], pv('Length')
            if L and PL and str(PL).isdigit() and int(L) > int(PL):
                pr.append('Length: VLM=%s > variable length %s -> used %s (DD0123)' % (L, PL, PL)); L = PL
            put('Length', str(L) if L else None)
            o = map_origin(r[h['origin']]); put('Origin', o); put('Source', map_source(r[h['origin']], tab, var))
            n_, o_ = txt(r[h['conversiondefinition']]), txt(r[h.get('programmingnotes', h.get('sascodes'))])
            t = n_ or o_
            if n_ and o_ and norm(n_) != norm(o_): pr.append('N/O differ - N used; O="%s"' % o_[:150])
            wid = '%s.%s.%s' % (tab, var, re.sub(r'[^A-Za-z0-9]+', '_', txt(r[h['wherevalue']]))[:40].strip('_'))
            if o == 'Derived' and t: methods['MT.' + wid] = t; put('Method', 'MT.' + wid)
            elif o == 'Assigned' and t: comments['COM.' + wid] = t; put('Comment', 'COM.' + wid)
            elif not t: pr.append('no Conversion Definition/Programming Notes in VLM'); VL.cell(nr, lci['Comment' if o == 'Assigned' else 'Method'] + 1).fill = ORG
            put_pr(VL, lci, nr, pr); log['VLM rows added'] += 1

# DD0074: variable-level Origin must equal every value-level Origin, else blank
# Reviewer convention: when EVERY value-level row of a variable has an Origin, the
# Variables row carries no origin info at all (Origin/Source/Pages/Method/Predecessor/
# Comment blank) - it is defined at value level only. Incomplete value level falls back
# to the DD0074 alignment below.
vlo, vln, vlfilled = defaultdict(set), Counter(), Counter()
for rr in range(2, VL.max_row + 1):
    k = (txt(VL.cell(rr, lci['Dataset'] + 1).value), txt(VL.cell(rr, lci['Variable'] + 1).value))
    if not k[0]: continue
    o = txt(VL.cell(rr, lci['Origin'] + 1).value); vln[k] += 1
    if o: vlo[k].add(o); vlfilled[k] += 1
for k, os_ in vlo.items():
    rr = vrow.get(k)
    if not rr or k[1] == 'QVAL': continue
    if vlfilled[k] == vln[k]:
        old = {n: V.cell(rr, vci[n] + 1).value for n in ('Origin', 'Source', 'Pages', 'Method', 'Predecessor', 'Comment')
               if n in vci and txt(V.cell(rr, vci[n] + 1).value)}
        for n in old: V.cell(rr, vci[n] + 1).value = None
        V.cell(rr, vci['Origin'] + 1).fill = NOFILL
        prc = V.cell(rr, vci['Programmer Review'] + 1)
        prc.value = re.sub(r'Origin: (skeleton=[^;]*|no source[^;]*);? ?', '', txt(prc.value)) or None
        put_pr(V, vci, rr, ['Defined at value level only (%d rows, origins %s)%s' % (vln[k], sorted(os_),
               '; cleared ' + ', '.join('%s=%s' % kv for kv in old.items()) if old else '')])
        log['var-level cleared (VLM only)'] += 1
        continue
    c = V.cell(rr, vci['Origin'] + 1); cur = txt(c.value)
    d = specvars.get(k); so = map_origin(d['r_orig'] or d['origin']) if d else ''
    prc = V.cell(rr, vci['Programmer Review'] + 1)
    prc.value = re.sub(r'Origin: (skeleton=[^;]*|no source[^;]*);? ?', '', txt(prc.value)) or None
    if len(os_) == 1 and next(iter(os_)) in (cur, so):
        new = next(iter(os_))
        if cur != new: put_pr(V, vci, rr, ['Origin: skeleton=%s -> %s (spec and all value-level origins agree, DD0074)' % (cur or 'blank', new)])
        c.value = new; c.fill = NOFILL; log['DD0074 aligned'] += 1
    elif len(os_) == 1 and not cur:
        c.value = next(iter(os_)); c.fill = NOFILL
    else:
        c.value = None; c.fill = NOFILL
        for n in ('Source', 'Pages', 'Method', 'Comment'): V.cell(rr, vci[n] + 1).value = None
        put_pr(V, vci, rr, ['Origin defined at value level only %s (DD0074)' % sorted(os_)]); log['DD0074 cleared'] += 1
    if txt(c.value): V.cell(rr, vci['Source'] + 1).value = V.cell(rr, vci['Source'] + 1).value or ('Sponsor' if c.value != 'Collected' else None)

# ---------------- Datasets / Standards ----------------
D = wb['Datasets']; dci = add_pr_col(D)
present = set()
for rr in range(2, D.max_row + 1):
    ds = txt(D.cell(rr, dci['Dataset'] + 1).value)
    if not ds: continue
    present.add(ds); pr = []
    if IGVER and 'Standard' in dci and not txt(D.cell(rr, dci['Standard'] + 1).value):
        D.cell(rr, dci['Standard'] + 1).value = 'SDTMIG ' + IGVER
    cr = content.get(ds)
    if not ds.startswith('SUPP'):
        if ds not in in_scope: pr.append('not marked for %s on Content tab' % STUDY)
        if cr:
            for col, key in (('Class', 'class'), ('Structure', 'structure'), ('Key Variables', 'sortorder')):
                c = D.cell(rr, dci[col] + 1)
                if key in ch and (not txt(c.value) or (col == 'Class' and txt(c.value) not in VALIDCLASS)):
                    new = txt(cr[ch[key]])
                    if col == 'Class' and new not in VALIDCLASS:
                        c.fill = ORG; pr.append('Class: skeleton=%s, Content tab=%s - neither is a valid Define-XML class (DD0055) - decide' % (c.value, new))
                        continue
                    if col == 'Key Variables':   # P21 DD0040 needs a comma-separated list
                        new = ', '.join(t for t in re.split(r'[\s,]+', new) if t)
                    pr.append('%s: skeleton=%s -> Content tab' % (col, c.value or 'blank')); c.value = new
    put_pr(D, dci, rr, pr)
for ds in sorted(in_scope - present): review.append('Content tab marks %s for %s but it is not in the skeleton (no data?)' % (ds, STUDY))
if 'Standards' not in wb.sheetnames and IGVER and CTVER:
    S = wb.create_sheet('Standards')
    S.append(['Name', 'Type', 'Publishing Set', 'Version', 'Status', 'Comment'])
    S.append(['SDTMIG', 'IG', None, IGVER, 'Final', None]); S.append(['SDTM', 'CT', 'SDTM', CTVER, 'Final', None])

# ---------------- Codelists: Terminology (DD0127) and unique Names (DD0131) ----------------
CL = wb['Codelists']; cli = {c.value: i for i, c in enumerate(CL[1])}
users = defaultdict(set)
for ws, ci in ((V, vci), (VL, lci)):
    for rr in range(2, ws.max_row + 1):
        c = txt(ws.cell(rr, ci['Codelist'] + 1).value)
        if c: users[c].add(txt(ws.cell(rr, ci['Dataset'] + 1).value))
names = {}
for rr in range(2, CL.max_row + 1):
    i = txt(CL.cell(rr, 1).value)
    if i: names.setdefault(i, txt(CL.cell(rr, cli['Name'] + 1).value))
dup = Counter(names.values())
for rr in range(2, CL.max_row + 1):
    i = txt(CL.cell(rr, 1).value)
    if not i: continue
    if CTVER and txt(CL.cell(rr, cli['NCI Codelist Code'] + 1).value) and not txt(CL.cell(rr, cli['Terminology'] + 1).value):
        CL.cell(rr, cli['Terminology'] + 1).value = 'SDTM ' + CTVER
    if dup[names[i]] > 1:
        m = re.match(r'QVAL\.(SUPP[A-Z0-9]+)\.', i)
        m2 = re.match(r'([A-Z][A-Z0-9]{1,7})\.([A-Z][A-Z0-9]{1,7})\.', i)   # value-level ID VAR.DS.WHERE...
        dom = m.group(1) if m else (m2.group(1) if m2 else (next(iter(users[i])) if len(users[i]) == 1 else i[:2]))
        CL.cell(rr, cli['Name'] + 1).value = '%s (%s)' % (names[i], dom)

# ---------------- Methods / Comments ----------------
M = wb['Methods']; mci = {c.value: i for i, c in enumerate(M[1])}
for mid in sorted(methods):
    r = [None] * len(mci); r[mci['ID']] = mid; r[mci['Name']] = mid; r[mci['Type']] = 'Computation'; r[mci['Description']] = methods[mid]; M.append(r)
C = wb['Comments']; cci = {c.value: i for i, c in enumerate(C[1])}
for cid in sorted(comments):
    r = [None] * len(cci); r[cci['ID']] = cid; r[cci['Description']] = comments[cid]; C.append(r)

out = args.out
try: wb.save(out)
except PermissionError:
    out = out.replace('.xlsx', '_updated.xlsx'); wb.save(out)
openpyxl.load_workbook(out)   # must re-open cleanly
print('saved', out, '| study', STUDY, '| SDTMIG', IGVER, '| CT', CTVER)
for k, v in sorted(log.items()): print('  %-28s %d' % (k, v))
print('  methods %d, comments %d' % (len(methods), len(comments)))
for x in review: print('REVIEW', x)
