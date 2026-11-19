import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# A4, docDefaults Calibri 11, useFELayout. Japanese body paragraphs whose runs pair a
# Latin ascii face with an MS Gothic / MS Mincho / MS PGothic eastAsia face: Word paces
# every pair at the East Asian factor (1.3029 single, 1.2952-1.30 under line=288) whatever
# the Latin partner (Cambria / Calibri / Times New Roman / Arial), probe 2026-11-07.
A4 = '<w:pgSz w:w="11900" w:h="16840"/><w:pgMar w:top="1080" w:right="1152" w:bottom="1656" w:left="1152" w:header="720" w:footer="864" w:gutter="0"/>'
JP = '日本語の段落テキスト行の高さを測る'
MIX = '日本語 abc Word 123 テキスト'
def block(ascii, ea, sz, line):
    return para([JP, MIX, JP, MIX], spacing(after=0, line=line, rule='auto'), rpr(ascii, sz, ea))
body = (
    block('Cambria', 'MS Gothic', 24, 288)
    + block('Calibri', 'MS Gothic', 24, 288)
    + block('Times New Roman', 'MS Mincho', 21, 240)
    + block('Arial', 'MS PGothic', 24, 240)
    + block('MS Gothic', 'MS Gothic', 24, 288)
)
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml('<w:useFELayout/>'), sect=A4)
