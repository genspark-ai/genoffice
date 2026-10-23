import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
AB = rpr('Arial', 22, extra='<w:b/>')
body = para('Phone: 012 345 6789\tEmail: x@y.z', run_rpr=AB) + para('CORE FOCUS\t\t\t:', run_rpr=AB)
write(out_path(), body, styles_xml(doc_defaults('Arial', 22)), settings_xml())
