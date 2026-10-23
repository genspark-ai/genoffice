import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
T = rpr('Times New Roman', 24)
L1 = 'Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore et dolore magna aliqua ut enim ad minim veniam quis nostrud'
L3 = 'Masse salariale chargee du personnel'
def tbl(grid, texts, pct=5000):
    g = ''.join(f'<w:gridCol w:w="{w}"/>' for w in grid)
    cells = ''.join(f'<w:tc><w:tcPr/>{para(t, run_rpr=T)}</w:tc>' for t in texts)
    return f'<w:tbl><w:tblPr><w:tblW w:w="{pct}" w:type="pct"/></w:tblPr><w:tblGrid>{g}</w:tblGrid><w:tr>{cells}</w:tr></w:tbl>'
body = para('A') + tbl([3000, 3000], [L1, L3]) + para('B') + tbl([1000, 8026], [L1, L3]) + para('C')
tn = '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
write(out_path(), body, styles_xml(doc_defaults('Times New Roman', 24), tn), settings_xml(), sect=A4)
