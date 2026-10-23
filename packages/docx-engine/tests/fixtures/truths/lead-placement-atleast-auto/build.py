import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
T = rpr('Times New Roman', 28)
body = (para('TNR 14 line 360 auto', ppr=spacing(after=0, line=360, rule='auto'), run_rpr=T)
        + para('TNR 14 exact 480', ppr=spacing(after=0, line=480, rule='exact'), run_rpr=T)
        + para('TNR 14 atLeast 480', ppr=spacing(after=0, line=480, rule='atLeast'), run_rpr=T))
write(out_path(), body, styles_xml(doc_defaults('Times New Roman', 24)), settings_xml(), sect=A4)
