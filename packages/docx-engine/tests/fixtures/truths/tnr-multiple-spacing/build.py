import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
tnr12 = rpr('Times New Roman', 24)
tnr11 = rpr('Times New Roman', 22)
body = ''.join([
    para(['single line one', 'line two', 'line three'], run_rpr=tnr12),
    para(['line 360 one', 'line two', 'line three'], ppr=spacing(after=0, line=360, rule='auto'), run_rpr=tnr12),
    para(['11pt line 276 one', 'line two', 'line three'], ppr=spacing(after=0, line=276, rule='auto'), run_rpr=tnr11),
    para(['exact 300 one', 'line two'], ppr=spacing(after=0, line=300, rule='exact'), run_rpr=tnr12),
    para(['atLeast 400 one', 'line two'], ppr=spacing(after=0, line=400, rule='atLeast'), run_rpr=tnr12),
    para(['atLeast 200 floor below natural'], ppr=spacing(after=0, line=200, rule='atLeast'), run_rpr=tnr12),
])
write(out_path(), body, styles_xml(doc_defaults('Times New Roman', 24)), settings_xml())
