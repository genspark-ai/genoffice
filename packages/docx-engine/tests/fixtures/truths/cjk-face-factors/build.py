import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# East Asian faces named in Latin letters only (fixture hygiene); the sample text is hangul / kana / Han
HANGUL = '한글 줄 하나'
KANA = 'あいうえお カタカナ'
HAN = '文字段落行高'
def ea(font, sz, text):
    return para([text, text], run_rpr=rpr('Calibri', sz, east_asia=font))
body = ''.join([
    ea('Batang', 22, HANGUL),
    ea('Malgun Gothic', 22, HANGUL),
    ea('Meiryo', 21, KANA),
    ea('Yu Mincho', 21, KANA),
    ea('Microsoft YaHei', 24, HAN),
    ea('SimSun', 24, HAN),
    ea('MS Mincho', 21, KANA),
    ea('PMingLiU', 24, HAN),
])
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml('<w:useFELayout/>'))
