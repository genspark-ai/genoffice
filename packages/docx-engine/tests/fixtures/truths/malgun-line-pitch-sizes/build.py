import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
MG = '맑은 고딕'
H = '참조 문서 계약 범위'
def p(sz, sp):
    return para(H, ppr=sp, run_rpr=rpr('Calibri', sz, east_asia=MG))
body = ''.join([
    p(17, spacing(after=0, line=259, rule='auto')), p(18, spacing(after=0, line=259, rule='auto')), p(19, spacing(after=0, line=259, rule='auto')),
    p(20, spacing(after=0, line=259, rule='auto')), p(22, spacing(after=0, line=259, rule='auto')),
    p(18, spacing(after=0, line=240, rule='auto')), p(18, spacing(after=0, line=276, rule='auto')), p(18, spacing(after=0, line=288, rule='auto')),
    p(18, spacing(after=0, line=360, rule='exact')), p(18, spacing(after=0, line=240, rule='atLeast')),
])
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22, east_asia=MG, ppr=spacing(after=160, line=259, rule='auto'))), settings_xml('<w:useFELayout/>'), sect=A4)
