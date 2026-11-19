import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
B = '<w:tblBorders>' + ''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>' for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')) + '</w:tblBorders>'
def tbl(grid, dxa, jc=None, nested=None):
    g = ''.join(f'<w:gridCol w:w="{w}"/>' for w in grid)
    j = f'<w:jc w:val="{jc}"/>' if jc else ''
    cells = ''
    for i, w in enumerate(grid):
        inner = para(f'c{i}')
        if nested is not None and i == nested[0]:
            inner = nested[1] + para('')
        cells += f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>{inner}</w:tc>'
    return f'<w:tbl><w:tblPr><w:tblW w:w="{dxa}" w:type="dxa"/>{j}{B}</w:tblPr><w:tblGrid>{g}</w:tblGrid><w:tr>{cells}</w:tr></w:tbl>'
inner = tbl([3000, 3000], 6000)
body = (para('A') + tbl([3000, 6000, 3000], 12000, 'center')
        + para('B') + tbl([3250, 6500, 3250], 13000, 'center')
        + para('C') + tbl([3000, 6000, 3000], 12000, 'right')
        + para('D') + tbl([2000, 5000, 3000], 10000, 'center', nested=(1, inner))
        + para('E') + tbl([3000, 6000, 3000], 12000)
        + para('F'))
tn = '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
write(out_path(), body, styles_xml(doc_defaults(), tn), settings_xml())
