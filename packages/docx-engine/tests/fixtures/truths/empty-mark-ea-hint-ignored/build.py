import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
B = 'Batang'
SP = spacing(after=0, line=240, rule='auto')
def empty(mark):
    return f'<w:p><w:pPr>{SP}<w:rPr>{mark}</w:rPr></w:pPr></w:p>'
body = (empty('<w:rFonts w:hint="eastAsia"/>') + empty('<w:lang w:eastAsia="ko-KR"/>') + empty(f'<w:rFonts w:eastAsia="{B}"/>')
        + empty(f'<w:rFonts w:eastAsia="{B}" w:hint="eastAsia"/>') + empty(f'<w:rFonts w:ascii="{B}" w:hAnsi="{B}"/>')
        + para('한글 text in Noto Sans CJK SC', run_rpr=rpr('Noto Sans CJK SC', 22, east_asia='Noto Sans CJK SC')))
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml(), sect=A4)
