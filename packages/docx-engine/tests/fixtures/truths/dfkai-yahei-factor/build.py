import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
body = para('文字 DFKai', run_rpr=rpr('Calibri', 24, east_asia='DFKai-SB'))
write(out_path(), body, styles_xml(doc_defaults('Calibri', 24, east_asia='DFKai-SB')), settings_xml('<w:useFELayout/>'))
