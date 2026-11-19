import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Exact and atLeast rules inside cells: exact 24pt pins two 11pt lines at 48pt,
# exact 10pt clips the line to 10pt, atLeast 20pt lays two lines at 20pt each.
two = ['first line', 'second line']
rows = [
    para(two, ppr=spacing(line=480, rule='exact')),
    para(two, ppr=spacing(line=200, rule='exact')),
    para(two, ppr=spacing(line=400, rule='atLeast')),
    para(two),
]
body = para('lead') + tbl(rows) + para('tail')
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
