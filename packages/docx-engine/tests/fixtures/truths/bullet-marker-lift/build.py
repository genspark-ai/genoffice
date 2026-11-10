import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
body = para('P0 plain line', run_rpr=rpr('Calibri', 22))
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
