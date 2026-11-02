import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Barlow Condensed (M365 cloud font) lines at 11 pt, upright and bold, plus the
# weight-worded spelling Word maps to the same family; a Calibri control line.
TEXT = 'The quick brown fox jumps over the lazy dog 0123456789'
BODY = 'On the Insert tab, the galleries include items that are designed to coordinate with the overall look of your document.'
body = para('control')
body += para(TEXT, run_rpr=rpr('Barlow Condensed', 22))
body += para(TEXT, run_rpr=rpr('Barlow Condensed', 22, extra='<w:b/>'))
body += para(BODY, run_rpr=rpr('Barlow Condensed', 22))
body += para(TEXT, run_rpr=rpr('Barlow Condensed Bold', 22))
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
