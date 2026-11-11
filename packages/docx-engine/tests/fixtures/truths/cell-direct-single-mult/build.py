import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Normal carries line 276 (1.15); a cell paragraph with a direct line 240 auto is
# single-spaced (Calibri 11 pitch 13.43), its neighbour inherits the 1.15 (15.44).
lines = ['line one', 'line two', 'line three']
rows = [
    para(lines, ppr=spacing(line=240, rule='auto')),
    para(lines),
]
body = para('lead') + tbl(rows) + para('tail')
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22, ppr=spacing(after=0, line=276, rule='auto'))), settings_xml())
