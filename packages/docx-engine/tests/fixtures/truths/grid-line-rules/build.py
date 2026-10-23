import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
cal = rpr('Calibri', 22)
body = ''.join([
    para('single on an 18pt grid', run_rpr=cal),
    para('line 360 auto scales the pitch', ppr=spacing(after=0, line=360, rule='auto'), run_rpr=cal),
    para('exact 300 never snaps', ppr=spacing(after=0, line=300, rule='exact'), run_rpr=cal),
    para('atLeast 200 floors at the cell', ppr=spacing(after=0, line=200, rule='atLeast'), run_rpr=cal),
    para('atLeast 0 keeps the natural height', ppr=spacing(after=0, line=0, rule='atLeast'), run_rpr=cal),
    para('16pt single takes two cells', run_rpr=rpr('Calibri', 32)),
    para('snapToGrid off leaves the grid', ppr='<w:snapToGrid w:val="0"/>' + spacing(after=0, line=240, rule='auto'), run_rpr=cal),
    para('line 480 auto is two pitches', ppr=spacing(after=0, line=480, rule='auto'), run_rpr=cal),
])
sect = A4 + '<w:docGrid w:type="lines" w:linePitch="360"/>'
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml(), sect=sect)
