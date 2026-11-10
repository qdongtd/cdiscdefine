# -*- coding: utf-8 -*-
"""Fill a P21-generated SDTM or ADaM Define-XML 2.1 spec workbook (entry point).

Detects the standard from the skeleton's Datasets classes (or use --standard) and
runs the matching engine with the remaining arguments:

  SDTM -> sdtm_fill.py : --skeleton --spec --acrf --out [--vlm] [--study] [--sibling]
          [--page-ref physical|printed] [--subject-domains QS] [--ct] [--domain-vl inherit|flag|acrf]
          [--vendor-domains LB,LC] [--nodata-text]
  ADaM -> adam_fill.py : --skeleton --spec --out [--vlm] [--issues] [--study] [--ct]
          [--keep-p21-dataset-meta] [--nodata-text]

    python fill_define.py [--standard sdtm|adam] --skeleton <p21_define.xlsx> ...
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
if std is None:
    if '--skeleton' not in argv: sys.exit('fill_define.py: --skeleton is required')
    std = detect_standard(openpyxl.load_workbook(argv[argv.index('--skeleton') + 1]))
if std not in ('sdtm', 'adam'): sys.exit('--standard must be sdtm or adam')
print('standard:', std.upper())
engine = os.path.join(HERE, 'sdtm_fill.py' if std == 'sdtm' else 'adam_fill.py')
sys.argv = [engine] + argv
runpy.run_path(engine, run_name='__main__')
