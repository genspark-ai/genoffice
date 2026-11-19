import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *


def tbl(tcws, rows):
    g = ''.join(f'<w:gridCol w:w="{w}"/>' for w in tcws)
    trs = ''.join(
        '<w:tr>' + ''.join(
            f'<w:tc><w:tcPr><w:tcW w:w="{tcws[i]}" w:type="dxa"/></w:tcPr>{para(t)}</w:tc>'
            for i, t in enumerate(r)) + '</w:tr>'
        for r in rows)
    return f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>{g}</w:tblGrid>{trs}</w:tbl>'


narrow = lambda text: ([600, 8400], [[text, 'filler text that wraps across a few lines so the wide column has surplus']])
body = (para('A') + tbl(*narrow('mmmm–mmmm'))
        + para('B') + tbl(*narrow('mmmm—mmmm'))
        + para('C') + tbl(*narrow('mmmm−mmmm'))
        + para('D'))
A4 = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
write(out_path(), body, styles_xml(doc_defaults()), settings_xml(), sect=A4)
