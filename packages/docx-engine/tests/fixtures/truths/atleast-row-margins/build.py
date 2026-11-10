import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
def bd(sz):
    return ''.join(f'<w:{s} w:val="single" w:sz="{sz}" w:space="0" w:color="auto"/>' for s in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'])
def tbl(label, trh, rule, mar_top, mar_bottom, n=3):
    rule_attr = f' w:hRule="{rule}"' if rule else ''
    mars = f'<w:tblCellMar><w:top w:w="{mar_top}" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="{mar_bottom}" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
    rows = ''.join(
        f'<w:tr><w:trPr><w:trHeight{rule_attr} w:val="{trh}"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>{para(f"{label} r{i}")}</w:tc></w:tr>'
        for i in range(n)
    )
    return para(label) + f'<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblBorders>{bd(8)}</w:tblBorders><w:tblLayout w:type="fixed"/>{mars}</w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>{rows}</w:tbl>'
body = ''.join([
    tbl('A atLeast 440 mar 100/100', 440, 'atLeast', 100, 100),
    tbl('B atLeast 440 mar 0/0', 440, 'atLeast', 0, 0),
    tbl('C plain trHeight 440 mar 100/100', 440, None, 100, 100),
])
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
