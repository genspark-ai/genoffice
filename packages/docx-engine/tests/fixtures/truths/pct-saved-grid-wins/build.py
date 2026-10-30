import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
B = '<w:tblBorders>' + ''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>' for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')) + '</w:tblBorders>'
PCT = [750, 2000, 1500, 750]
def tbl(grid, pct=5000, tcw=PCT, tcw_type='pct'):
    g = ''.join(f'<w:gridCol w:w="{w}"/>' for w in grid)
    cells = ''.join(f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="{tcw_type}"/></w:tcPr>{para(f"c{i}")}</w:tc>' for i, w in enumerate(tcw))
    return f'<w:tbl><w:tblPr><w:tblW w:w="{pct}" w:type="pct"/>{B}</w:tblPr><w:tblGrid>{g}</w:tblGrid><w:tr>{cells}</w:tr></w:tbl>'
G = [1909, 3473, 2536, 1442]
body = (para('A') + tbl(G)
        + para('B') + tbl([2340, 2340, 2340, 2340])
        + para('C') + tbl([1224, 2226, 1626, 924])
        + para('D') + tbl(G, tcw=G, tcw_type='dxa')
        + para('E') + tbl([1145, 2084, 1522, 865], pct=3000)
        + para('F'))
tn = '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
write(out_path(), body, styles_xml(doc_defaults(), tn), settings_xml())
