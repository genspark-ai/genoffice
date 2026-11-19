import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Korean text in a narrow text box wraps at spaces (keep-all), never inside a word.
KO = '한글 단어는 공백에서만 줄바꿈 됩니다 긴단어들이연속되어도 마찬가지'
content = para(KO, run_rpr=rpr('Malgun Gothic', 22, east_asia='Malgun Gothic'))
body = para('lead') + textbox(content, cx=1828800, cy=1097280) + para('tail')
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22, east_asia='Malgun Gothic')), settings_xml())
