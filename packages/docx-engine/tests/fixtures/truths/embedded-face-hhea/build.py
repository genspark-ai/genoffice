import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
body = para('embedded probe', run_rpr=rpr('GoProbeA', 24))
write(out_path(), body, styles_xml(doc_defaults('Calibri', 24)), settings_xml())
