import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Bookman Old Style (Office-only, DFonts) lines at 11 pt in the four cuts; a
# Calibri control line. Word paces them at the font's own hmtx advances.
TEXT = 'The quick brown fox jumps over the lazy dog 0123456789'
BODY = 'On the Insert tab, the galleries include items that are designed to coordinate with the overall look of your document.'
cuts = [('', ''), ('<w:b/>', 'bold'), ('<w:i/>', 'italic'), ('<w:b/><w:i/>', 'bold italic')]
body = para('control')
for extra, _ in cuts:
    body += para(TEXT, run_rpr=rpr('Bookman Old Style', 22, extra=extra))
body += para(BODY, run_rpr=rpr('Bookman Old Style', 22))
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
