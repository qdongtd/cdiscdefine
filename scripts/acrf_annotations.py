# -*- coding: utf-8 -*-
"""
Extract SDTM variable -> CRF page numbers from an annotated CRF (aCRF) PDF,
using the annotation layer with NOTE-box exclusion, RELREC left-side counting,
and box-fill-color -> domain disambiguation.

This is a STARTING POINT, not a black box. aCRFs vary (annotation font, box
colors, header format). Read references/acrf-page-derivation.md, then run this,
inspect the printed per-variable page lists, and adjust the tunables at the top
for the specific aCRF before trusting the output. The aCRF is authoritative for
page VALUES: when it disagrees with the spec, change the spec, not the aCRF.

Usage:
    python acrf_annotations.py acrf.pdf [--var-domains var2dom.json] [--json out.json]

    var2dom.json (optional): {"AESPID": ["AE"], "AGE": ["DM"], ...}
      -> map each SDTM variable to the domain(s) whose box color legitimises it.
      If omitted, every Arial token is reported with the domain(s) inferred from
      the box it sits in, which is looser (no cross-domain filtering).

Requires: PyMuPDF  (pip install pymupdf)  ->  import fitz
"""
import sys, re, json, argparse
from collections import defaultdict

import fitz  # PyMuPDF

# ---- tunables ----------------------------------------------------------------
ANNOT_FONTS = ('ArialMT', 'Arial-BoldMT', 'Arial')  # annotation layer font(s)
NOTE_RE     = re.compile(r'^\s*note\s*[:\-]', re.I)  # NOTE box -> exclude
HEADER_RE   = re.compile(r'^([A-Z]{2})\s*\(')        # "AE (Adverse Events)" -> AE
RELREC_RE   = re.compile(r'\bRELREC\b', re.I)
VAR_RE      = re.compile(r'^[A-Z][A-Z0-9]{1,7}$')    # a bare SDTM variable token
# lowercase words allowed inside a structured (non-prose) annotation span
CONNECTORS  = {'in', 'when', 'and', 'or', 'eq', 'ne', 'to', 'per', 'of'}
COLOR_ROUND = 1                                       # round fill colors to N dp
CONTAIN_FRAC = 0.30                                   # box must cover >30% of span
# DM/DC share the demographics palette; extend as needed for a given study.
DOMAIN_EQUIV = {'DC': 'DM', 'DM': 'DM'}
# -----------------------------------------------------------------------------


RANGE_RE = re.compile(r'\b([A-Z][A-Z0-9]*?[A-Z])(\d+)\s+to\s+\1(\d+)\b')   # SVUPDE1 to SVUPDE12


def split_tokens(text):
    """Variable-name candidates in an annotation span.

    Splits on whitespace AND on '/', '=', ',', ';', ':', '(', ')' so that
    'DSTERM/DSDECOD=INFORMED CONSENT' yields DSTERM, DSDECOD (the value glued
    after '=' is dropped), and expands ranges 'SUDRINK1 to SUDRINK4' into
    SUDRINK1..SUDRINK4. A plain whitespace split misses every variable glued
    to '/' or '='.
    """
    out = []
    for w in text.split():
        w = w.split('=')[0]                      # VAR=value -> VAR
        for part in re.split(r'[/,;:()]', w):
            part = part.strip('.')
            if part:
                out.append(part)
    for stem, a, b in RANGE_RE.findall(text):
        if int(a) <= int(b) <= int(a) + 50:
            out.extend('%s%d' % (stem, n) for n in range(int(a), int(b) + 1))
    return out


def is_struct(text):
    """True if the span looks like a variable/structured annotation, not prose."""
    for w in re.findall(r'[A-Za-z]+', text):
        if w.islower() and w.lower() not in CONNECTORS:
            return False
    return bool(text.strip())


def round_color(c):
    if c is None:
        return None
    if isinstance(c, (int, float)):  # grayscale
        c = (c, c, c)
    return tuple(round(float(x), COLOR_ROUND) for x in c[:3])


def page_boxes(page):
    """Return [(fitz.Rect, rounded_fill_color)] for filled rectangles on page."""
    out = []
    for d in page.get_drawings():
        fill = d.get('fill')
        if fill is None:
            continue
        rect = d.get('rect')
        if rect is None:
            for it in d.get('items', []):
                if it[0] == 're':
                    rect = it[1]
                    break
        if rect is not None:
            out.append((fitz.Rect(rect), round_color(fill)))
    return out


def smallest_box_rect(boxes, span_rect):
    """(rect, fill color) of the smallest box that meaningfully contains the span."""
    best, area = (None, None), float('inf')
    for rect, col in boxes:
        if col is None:
            continue
        inter = (rect & span_rect)
        if inter.get_area() > CONTAIN_FRAC * max(span_rect.get_area(), 1e-6):
            a = rect.get_area()
            if a < area:
                area, best = a, (rect, col)
    return best


def smallest_box(boxes, span_rect):
    """Fill color of the smallest box that meaningfully contains the span."""
    return smallest_box_rect(boxes, span_rect)[1]


def _rkey(rect):
    return None if rect is None else tuple(round(v, 1) for v in rect)


def norm_dom(d):
    return DOMAIN_EQUIV.get(d, d)


def extract(pdf_path, offset=None):
    doc = fitz.open(pdf_path)
    if offset is None:
        offset = detect_cover_offset(doc)  # printed = pdf_0based - offset + 1
    # var -> set(pages), plus color->domain per page
    var_pages = defaultdict(set)
    pages = []                       # (crf_page, col2dom, spans, note_boxes) per page
    global_c2d = defaultdict(set)    # color -> domains seen in header boxes anywhere in the PDF
    for pno in range(doc.page_count):
        page = doc[pno]
        crf_page = pno - offset + 1  # printed CRF page (cover offset applied)
        if crf_page < 1:
            continue
        boxes = page_boxes(page)
        # 1) build color -> domain map from header spans on this page
        col2dom = {}
        spans = []
        note_boxes = []
        d = page.get_text('dict')
        for block in d.get('blocks', []):
            for line in block.get('lines', []):
                for sp in line.get('spans', []):
                    txt = sp['text'].strip()
                    if not txt:
                        continue
                    font = sp.get('font', '')
                    rect = fitz.Rect(sp['bbox'])
                    brect, col = smallest_box_rect(boxes, rect)
                    m = HEADER_RE.match(txt)
                    if m and col is not None:
                        col2dom.setdefault(col, m.group(1))
                    if NOTE_RE.match(txt) and brect is not None:
                        # mark THIS box as a NOTE box. Never exclude by color: a NOTE box
                        # often shares its domain's color, and excluding the color drops
                        # every genuine annotation of that domain on the page.
                        note_boxes.append(_rkey(brect))
                    if font.startswith(ANNOT_FONTS) or font in ANNOT_FONTS:
                        spans.append((txt, rect, col, _rkey(brect)))
        for c, dm in col2dom.items():
            global_c2d[c].add(norm_dom(dm))
        pages.append((crf_page, col2dom, spans, note_boxes))
    # A multi-page form usually carries the domain header box only on its first page;
    # continuation pages reuse the same box color. Fall back to the PDF-wide
    # color -> domain map when the color maps to exactly one domain.
    fallback = {c: next(iter(ds)) for c, ds in global_c2d.items() if len(ds) == 1}
    for crf_page, col2dom, spans, note_boxes in pages:
        # 2) walk annotation spans on this page
        for txt, rect, col, bkey in spans:
            if bkey is not None and bkey in note_boxes:   # filter 1: inside a NOTE box
                continue
            if not is_struct(txt):          # filter 2: prose
                continue
            toks = split_tokens(txt)
            if RELREC_RE.search(txt):       # filter 3: RELREC -> left var only
                # "<VAR> RELREC when X = Y" or "RELREC <VAR> when ..."
                cand = [t for t in toks if VAR_RE.match(t) and t != 'RELREC']
                toks = cand[:1]
            dom_here = norm_dom(col2dom.get(col) or fallback.get(col, '')) if col is not None else ''
            for t in toks:
                if VAR_RE.match(t):
                    var_pages[t].add((crf_page, dom_here))
    doc.close()
    return var_pages


def collapse(var_pages, var2dom=None):
    """Return {var: sorted_pages}, applying color->domain filter if var2dom given."""
    out = {}
    for var, entries in var_pages.items():
        if var2dom and var in var2dom:
            allowed = {norm_dom(x) for x in var2dom[var]}
            pages = sorted({p for (p, dom) in entries if dom in allowed})
            if not pages:  # color filter removed everything -> report unfiltered too
                pages = sorted({p for (p, _) in entries})
        else:
            pages = sorted({p for (p, _) in entries})
        out[var] = pages
    return out


def collapse_by_domain(var_pages):
    """Return {(var, domain): sorted_pages} keeping the box-domain of each page.

    Use this (not collapse) when a variable name appears in several domains
    (FOCID, --SPID, generic identifiers): assign pages to a Variables-tab ROW by
    the ROW's own dataset domain, so AE.FOCID gets only AE-colored pages and not
    the OE/MH/SC ones. Pages with an unresolved domain ('') are still returned
    under '' for a name-scan fallback.
    """
    out = defaultdict(set)
    for var, entries in var_pages.items():
        for (p, dom) in entries:
            out[(var, dom)].add(p)
    return {k: sorted(v) for k, v in out.items()}


PAGE_OF_RE = re.compile(r'\b(\d+)\s+of\s+(\d+)\b')


def detect_cover_offset(doc):
    """Printed CRF page numbers usually differ from PDF page indices by a cover.

    Find the PDF page that prints 'Page 1 of N' (or '1 of N') and return that
    page's 0-based index = the number of front-matter pages before printed page 1.
    Then: printed_page(pdf_0based p) = p - offset + 1. With the common single
    cover page, offset == 1 and printed == pdf_0based == (extractor 1-based page - 1).
    Returns 1 (assume one cover) if no 'N of N' footer is found -- verify by hand.
    """
    # Front matter can be long (a generated table of contents runs 20+ pages), so
    # scan the first half of the PDF, and test that front matter + N == page count
    # (allowing a few trailing pages), not that N alone is close to the page count.
    for pno in range(max(1, doc.page_count // 2)):
        for m in PAGE_OF_RE.finditer(doc[pno].get_text()):
            n = int(m.group(2))
            if m.group(1) == '1' and doc.page_count - 3 <= pno + n <= doc.page_count:
                return pno
    return 1


def scan_supp_qnam_pages(doc, offset=1):
    """SUPP qualifier pages via full-text '<QNAM> in SUPP<dom>' scan.

    More reliable than the color extractor for supplemental qualifiers, which the
    RELREC / multi-line-span filters often drop even though '<QNAM> in SUPPxx' is
    clearly annotated. Returns {(SUPPxx, QNAM): [printed pages]} using the same
    cover offset (printed = pdf_0based - offset + 1).
    """
    pat = re.compile(r'([A-Z][A-Z0-9]{1,7})\s+in\s+(SUPP[A-Z0-9]{2,4})')
    rng = re.compile(r'\b([A-Z][A-Z0-9]*?[A-Z])(\d+)\s+to\s+\1(\d+)\s+in\s+(SUPP[A-Z0-9]{2,4})')
    out = defaultdict(set)
    for p in range(doc.page_count):
        printed = p - offset + 1
        if printed < 1:
            continue
        text = re.sub(r'\s+', ' ', doc[p].get_text())   # ranges can wrap across lines
        for m in pat.finditer(text):
            out[(m.group(2), m.group(1))].add(printed)
        for stem, a, b, sp in rng.findall(text):          # 'SUDRINK1 to SUDRINK4 in SUPPSU'
            if int(a) <= int(b) <= int(a) + 50:
                for n in range(int(a), int(b) + 1):
                    out[(sp, '%s%d' % (stem, n))].add(printed)
    return {k: sorted(v) for k, v in out.items()}


def name_scan(doc, offset=1):
    """Loose fallback: {TOKEN: [printed pages]} for every variable-like token in the
    annotation-font layer (NOTE boxes still skipped), ignoring box color/domain.

    Use ONLY when the color-filtered extractor finds no page for a Collected
    variable, and flag the result for a light confirm (the domain is not verified).
    """
    out = defaultdict(set)
    for p in range(doc.page_count):
        printed = p - offset + 1
        if printed < 1:
            continue
        for b in doc[p].get_text('dict').get('blocks', []):
            for l in b.get('lines', []):
                for s in l.get('spans', []):
                    if s.get('font', '').startswith(ANNOT_FONTS) and not NOTE_RE.match(s['text']):
                        for t in set(split_tokens(s['text'])):
                            if VAR_RE.match(t):
                                out[t].add(printed)
    return {k: sorted(v) for k, v in out.items()}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pdf')
    ap.add_argument('--var-domains', help='JSON map var -> [domains]')
    ap.add_argument('--json', help='write result to this JSON file')
    ap.add_argument('--cover-offset', default='auto',
                    help="front-matter pages before printed page 1: 'auto' (default) or an int")
    args = ap.parse_args()

    var2dom = None
    if args.var_domains:
        with open(args.var_domains, encoding='utf-8') as f:
            var2dom = json.load(f)

    offset = None if args.cover_offset == 'auto' else int(args.cover_offset)
    result = collapse(extract(args.pdf, offset), var2dom)
    for var in sorted(result):
        pages = result[var]
        print('%-12s %s' % (var, ' '.join(map(str, pages)) or '(none)'))
    if args.json:
        with open(args.json, 'w', encoding='utf-8') as f:
            json.dump(result, f, indent=1)
        print('\nwrote', args.json)


if __name__ == '__main__':
    main()
