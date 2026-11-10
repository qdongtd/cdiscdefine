# -*- coding: utf-8 -*-
"""Self-QC pre-check of an SDTM or ADaM Define-XML 2.1 spec workbook (entry point).

Detects the standard from the Datasets classes (or use --standard) and runs the
matching rule set. A PRE-FILTER, not authoritative: still generate define.xml and run P21.

  SDTM -> sdtm_qc.py : <define.xlsx> [--acrf <acrf.pdf>] [--package <folder>]
  ADaM -> adam_qc.py : <define.xlsx>

    python self_qc.py [--standard sdtm|adam] <define.xlsx> [options]
"""
import os, sys, runpy, warnings
import openpyxl
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from common import detect_standard
warnings.simplefilter('ignore')

argv = sys.argv[1:]
std = None
if '--standard' in argv:
    i = argv.index('--standard'); std = argv[i + 1].lower(); del argv[i:i + 2]
book = next((a for a in argv if a.lower().endswith(('.xlsx', '.xlsm'))), None)
if not book: sys.exit('self_qc.py: give the define workbook (.xlsx)')
if std is None: std = detect_standard(openpyxl.load_workbook(book))
engine = os.path.join(HERE, 'sdtm_qc.py' if std == 'sdtm' else 'adam_qc.py')
sys.argv = [engine] + argv
runpy.run_path(engine, run_name='__main__')
