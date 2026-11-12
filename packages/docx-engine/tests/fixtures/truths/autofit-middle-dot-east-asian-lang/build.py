import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *

FILL = 'filler text that wraps across a few lines so the wide column has surplus'


def tbl(text, lang):
    rpr = f'<w:rPr><w:lang w:eastAsia="{lang}"/></w:rPr>' if lang else ''
    cells = ''.join(
        f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>{para(t, run_rpr=rpr)}</w:tc>'
        for w, t in ((600, text), (8400, FILL)))
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="600"/><w:gridCol w:w="8400"/></w:tblGrid>'
            f'<w:tr>{cells}</w:tr></w:tbl>')


DOT = 'mmmm · mmmm'
body = (para('A') + tbl(DOT, '')
        + para('B') + tbl(DOT, 'zh-CN')
        + para('C') + tbl(DOT, 'zh-TW')
        + para('D') + tbl(DOT, 'pt-BR')
        + para('E') + tbl(DOT, 'en-US')
        + para('F') + tbl('mmmm : mmmm', 'pt-BR')
        + para('G') + tbl('mmmm–mmmm', '')
        + para('G'))
write(out_path(), body, styles_xml(doc_defaults()), settings_xml(), sect=A4)
