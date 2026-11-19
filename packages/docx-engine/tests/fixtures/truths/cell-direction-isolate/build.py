import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# bidiVisual two-cell row: the LTR cell keeps "abc 123" in reading order while the
# column order mirrors; the RTL cell and an Arabic text-box paragraph lay out RTL.
AR = 'عنوان التقرير'
C12 = rpr('Calibri', 24)
RTL = rpr('Calibri', 24, extra='<w:rtl/>')
bord = ''.join(
    f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')
)
row = (
    '<w:tr>'
    f'<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>{para("abc 123", run_rpr=C12)}</w:tc>'
    f'<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>{para(AR, ppr="<w:bidi/>", run_rpr=RTL)}</w:tc>'
    '</w:tr>'
)
table = (
    '<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:bidiVisual/><w:tblLayout w:type="fixed"/>'
    f'<w:tblBorders>{bord}</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>{row}</w:tbl>'
)
box = textbox(para(AR, ppr='<w:bidi/>', run_rpr=RTL) + para('abc 123', run_rpr=C12), cx=2743200, cy=640080)
body = para('lead', run_rpr=C12) + table + para('mid', run_rpr=C12) + box + para('tail', run_rpr=C12)
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
