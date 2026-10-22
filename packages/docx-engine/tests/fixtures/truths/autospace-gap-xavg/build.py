import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
B = 'Batang'
body = para('Cambria 가나 12', run_rpr=rpr('Cambria', 24, east_asia=B)) + para('http://example.com/a/very/long/url', run_rpr=rpr('Noto Sans CJK KR', 22, east_asia='Noto Sans CJK KR'))
write(out_path(), body, styles_xml(doc_defaults('Calibri', 24)), settings_xml('<w:useFELayout/>', 14), sect=A4)
