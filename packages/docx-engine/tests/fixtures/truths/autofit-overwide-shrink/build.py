import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
def tbl(grid, tcws, ind=None):
    g = ''.join(f'<w:gridCol w:w="{w}"/>' for w in grid)
    indx = f'<w:tblInd w:w="{ind}" w:type="dxa"/>' if ind is not None else ''
    cells = ''.join(f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>{para(f"c{i}")}</w:tc>' for i, w in enumerate(tcws))
    return f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>{indx}</w:tblPr><w:tblGrid>{g}</w:tblGrid><w:tr>{cells}</w:tr></w:tbl>'
body = (para('A') + tbl([3600, 3600, 3600], [3600, 3600, 3600])
        + para('B') + tbl([3600, 3600, 3600], [3600, 3600, 3600], -108)
        + para('C') + tbl([3600, 3600, 3600], [3600, 3600, 3600], 558)
        + para('D') + tbl([2400, 6000, 3600], [2400, 6000, 3600])
        + para('E') + tbl([2000, 4000, 3000], [3600, 3600, 3600])
        + para('F'))
write(out_path(), body, styles_xml(doc_defaults()), settings_xml())
