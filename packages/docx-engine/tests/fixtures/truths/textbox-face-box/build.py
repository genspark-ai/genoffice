import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Text-box lines: a Calibri 11 paragraph's first baseline sits one hhea ascent
# (0.952 em) under the top inset; a SimSun 12 paragraph lays a 1.30 line.
CJK = '中文段落文本'
box_a = textbox(para(['Calibri eleven line', 'second line'], run_rpr=rpr('Calibri', 22)), cx=2743200, cy=548640, id=1)
box_b = textbox(para([CJK, CJK], run_rpr=rpr('SimSun', 24, east_asia='SimSun')), cx=2743200, cy=548640, id=2)
body = para('lead') + box_a + para('mid') + box_b + para('tail')
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
