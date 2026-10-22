import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
MU = 'Meiryo UI'
def p(sz, sp):
    return para('あいうえお abc', ppr=sp, run_rpr=rpr(MU, sz, east_asia=MU))
body = ''.join([
    p(20, spacing(after=0, line=0, rule='atLeast')),
    p(20, spacing(after=0, line=200, rule='atLeast')),
    p(20, spacing(after=0, line=330, rule='atLeast')),
    p(20, spacing(after=0, line=400, rule='atLeast')),
    p(20, spacing(after=0)),
    p(20, spacing(after=0, line=360, rule='auto')),
    p(16, spacing(after=0, line=0, rule='atLeast')),
    p(28, spacing(after=0, line=0, rule='atLeast')),
])
sect = A4 + '<w:docGrid w:type="lines" w:linePitch="350"/>'
write(out_path(), body, styles_xml(doc_defaults(MU, 20, east_asia=MU)), settings_xml('<w:useFELayout/>'), sect=sect)
