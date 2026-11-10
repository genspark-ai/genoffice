import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
body = ''.join([
    para('control after 200 line 240', ppr=spacing(before=0, after=200, line=240, rule='auto')),
    para('fractional after 199.9 before 60.4 line 275.5', ppr=spacing(before='60.4', after='199.9', line='275.5', rule='auto')),
    para('fractional atLeast 300.7', ppr=spacing(after=0, line='300.7', rule='atLeast')),
    para('fractional exact 275.9', ppr=spacing(after=0, line='275.9', rule='exact')),
])
write(out_path(), body, styles_xml(doc_defaults('Arial', 22)), settings_xml())
