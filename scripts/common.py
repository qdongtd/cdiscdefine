# -*- coding: utf-8 -*-
"""Helpers shared by the SDTM and ADaM define scripts (one copy, used by both)."""
import re, copy
from openpyxl.styles import PatternFill

ORG = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')   # needs human review
NOFILL = PatternFill(fill_type=None)
DATE = {'date', 'datetime', 'time', 'partialDate', 'partialTime', 'partialDatetime',
        'incompleteDatetime', 'durationDatetime', 'intervalDatetime'}
SDTM_CLASSES = {'SPECIAL PURPOSE', 'FINDINGS', 'FINDINGS ABOUT', 'EVENTS', 'INTERVENTIONS', 'TRIAL DESIGN',
                'RELATIONSHIP', 'STUDY REFERENCE'}
ADAM_CLASSES = {'SUBJECT LEVEL ANALYSIS DATASET', 'BASIC DATA STRUCTURE', 'OCCURRENCE DATA STRUCTURE', 'ADAM OTHER'}


def load_values(path, blank_stop=2000):
    """Values-only copy of a workbook for read-only checks (QC, review).

    Workbooks saved from Excel are often formatted down to row 1,048,576; a full
    openpyxl load of such a file takes minutes or hangs. This streams each sheet in
    read-only mode, keeps rows that have content, stops after `blank_stop`
    consecutive empty rows, and returns (in-memory workbook, {sheet: (rows kept,
    used-range rows)}) so the caller can report oversized sheets."""
    import openpyxl
    src = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = openpyxl.Workbook(); out.remove(out.active); sizes = {}
    for name in src.sheetnames:
        ws, dst, blank, kept = src[name], out.create_sheet(name), 0, 0
        for row in ws.iter_rows(values_only=True):
            if any(v not in (None, '') for v in row):
                dst.append(list(row)); kept += 1; blank = 0
            else:
                blank += 1
                if blank > blank_stop: break
        sizes[name] = (kept, ws.max_row or kept)
    # P21-exported workbooks carry no sheet dimensions and read EMPTY in read-only
    # mode; they are never Excel-padded, so a normal load is fast for them.
    core = [n for n in ('Datasets', 'Variables') if n in sizes]
    if core and all(sizes[n][0] <= 1 for n in core):
        full = openpyxl.load_workbook(path, data_only=True)
        return full, {n: (full[n].max_row, full[n].max_row) for n in full.sheetnames}
    return out, sizes


def strip_padding(src, dst):
    """Remove empty styled rows below the last row with a value, in every worksheet,
    by editing the sheet XML (seconds; keeps every real row and its formatting).

    The usual cause of a define that "has a million rows" is ONE stray formatted
    empty row near the bottom plus a <dimension> tag pointing at it; Excel itself is
    very slow to delete it. Returns {sheet xml: removed row numbers}."""
    import zipfile, shutil, os
    zin = zipfile.ZipFile(src); tmp = dst + '.tmp'; removed = {}
    zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    row_re = re.compile(r'<row [^>]*?(?:/>|>.*?</row>)', re.S)
    for it in zin.infolist():
        data = zin.read(it.filename)
        if it.filename.startswith('xl/worksheets/sheet') and it.filename.endswith('.xml'):
            x = data.decode('utf8')
            rows = row_re.findall(x)
            num = lambda s: int(re.search(r'\br="(\d+)"', s).group(1))
            filled = [num(s) for s in rows if '<v>' in s or '<is>' in s]
            last = max(filled) if filled else 1
            gone = [num(s) for s in rows if num(s) > last and '<v>' not in s and '<is>' not in s]
            if gone:
                x = row_re.sub(lambda m: '' if num(m.group(0)) in gone else m.group(0), x)
                x = re.sub(r'<dimension ref="([A-Z]+\d+):([A-Z]+)\d+"/>',
                           lambda m: '<dimension ref="%s:%s%d"/>' % (m.group(1), m.group(2), last), x)
                removed[it.filename] = gone; data = x.encode('utf8')
        zout.writestr(it, data)
    zout.close(); zin.close(); shutil.move(tmp, dst)
    return removed


def oversized(sizes, factor=10, floor=5000):
    """Sheets whose used range is far beyond their content (formatted empty rows)."""
    return {n: s for n, s in sizes.items() if s[1] > max(floor, factor * max(s[0], 1))}


def norm(s): return re.sub(r'[^0-9a-z]', '', str(s or '').lower())


def txt(x):
    s = '' if x is None else str(x).strip()
    return '' if s == 'None' else s


def hdr(row): return {norm(x): i for i, x in enumerate(row) if x is not None}


def red(c):
    f = copy.copy(c.font); f.color = 'FFFF0000'; c.font = f


def add_pr_col(ws):
    """Add the non-standard 'Programmer Review' column (P21 ignores it); return name->index."""
    h = [c.value for c in ws[1]]
    if 'Programmer Review' not in h:
        ws.cell(row=1, column=len(h) + 1, value='Programmer Review'); h.append('Programmer Review')
    return {n: i for i, n in enumerate(h)}


def put_pr(ws, ci, row, msgs):
    if not msgs: return
    c = ws.cell(row=row, column=ci['Programmer Review'] + 1)
    c.value = '; '.join(([txt(c.value)] if txt(c.value) else []) + msgs)


def detect_standard(wb):
    """"adam" or "sdtm" from a define workbook: dataset classes first, then dataset names.
    Open the workbook normally (not read_only): P21 exports lack sheet dimensions and
    read as empty in read-only mode."""
    ws = wb['Datasets']; h = [c.value for c in ws[1]]
    ci = h.index('Class') if 'Class' in h else None
    names, classes = set(), set()
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or not r[0]: continue
        names.add(str(r[0]).upper())
        if ci is not None and r[ci]: classes.add(str(r[ci]).strip().upper())
    if classes & ADAM_CLASSES: return 'adam'
    if classes & SDTM_CLASSES: return 'sdtm'
    return 'adam' if 'ADSL' in names or (names and all(n.startswith('AD') for n in names)) else 'sdtm'
