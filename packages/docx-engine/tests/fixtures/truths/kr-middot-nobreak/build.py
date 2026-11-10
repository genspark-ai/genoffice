import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
MG = '맑은 고딕'
body = para('접수·검토·보완·회신,', run_rpr=rpr('Calibri', 16, east_asia=MG)) + para('Page 1 of 2', ppr=spacing(after=0, line=259, rule='auto'), run_rpr=rpr('Calibri', 22, east_asia=MG))
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22, east_asia=MG, ppr=spacing(after=0, line=259, rule='auto'))), settings_xml('<w:useFELayout/>'), sect=A4)
