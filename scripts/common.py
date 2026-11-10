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
