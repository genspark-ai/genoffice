import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# a Latin-named Korean face declared in every slot with a charset-81 fontTable entry (installed on the probe machine)
FACE = 'Pretendard'
HANGUL = '한글 줄 하나 둘 셋'
LATIN = 'Hamburgefonstiv 123'
def p(text, line=240):
    return para([text, text], ppr=spacing(after=0, line=line, rule='auto'), run_rpr=rpr(FACE, 22, east_asia=FACE))
body = ''.join([p(HANGUL), p(LATIN), p(HANGUL, 276), p(HANGUL, 360)])
fonts = font_table_xml([(FACE, {'panose1': '02000503000000020004', 'charset': '81', 'family': 'auto', 'pitch': 'variable'})])
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml('<w:useFELayout/>'), font_table=fonts, sect=A4)
