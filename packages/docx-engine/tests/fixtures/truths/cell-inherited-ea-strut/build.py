import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# CJK cell paragraphs inheriting the docDefaults East Asian face (SimSun) next to
# Latin ones inheriting Calibri: each row is its own face's single line.
CJK = '中文段落文本'
rows = [
    para(CJK),
    para('Latin control line'),
    para(CJK, run_rpr=rpr('Times New Roman')),
]
body = para('lead') + tbl(rows) + para('tail')
write(out_path(), body, styles_xml(doc_defaults('Calibri', 24, east_asia='SimSun')), settings_xml())
