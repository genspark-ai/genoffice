import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
def tbl(grid, tcws):
    g = ''.join(f'<w:gridCol w:w="{w}"/>' for w in grid)
    cells = ''.join(f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>{para(f"c{i}")}</w:tc>' for i, w in enumerate(tcws))
    return f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>{g}</w:tblGrid><w:tr>{cells}</w:tr></w:tbl>'
body = para('A') + tbl([2000, 2000, 2000], [3000, 3000, 3000]) + para('B') + tbl([10469], [6480]) + para('C')
write(out_path(), body, styles_xml(doc_defaults()), settings_xml())
