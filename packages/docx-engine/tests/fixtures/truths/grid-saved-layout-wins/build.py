import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
def tbl(grid, tcws, tblw):
    g = ''.join(f'<w:gridCol w:w="{w}"/>' for w in grid)
    cells = ''.join(f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>{para(t)}</w:tc>' for w, t in zip(tcws, ['a', 'b', 'Supercalifragilisticexpiali']))
    return f'<w:tbl><w:tblPr><w:tblW w:w="{tblw}" w:type="dxa"/></w:tblPr><w:tblGrid>{g}</w:tblGrid><w:tr>{cells}</w:tr></w:tbl>'
body = para('A') + tbl([966, 5188, 2846], [1000, 6500, 1500], 9000) + para('B')
write(out_path(), body, styles_xml(doc_defaults()), settings_xml())
