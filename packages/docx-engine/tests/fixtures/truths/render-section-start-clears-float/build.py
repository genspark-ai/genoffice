import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Letter, Calibri 11 single, two nextPage sections. Section 1: table T1 at the body top,
# a paragraph, a tblpPr table floated to the margin bottom (tblpYSpec bottom) and its
# anchor paragraph. Section 2 opens with the identical table T2: it must sit at the body
# top of page 2, not under the float the previous section left at its page bottom.
C = rpr('Calibri', 22)
SP = spacing(after=0, line=240, rule='auto')
BORD = ''.join(
    f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')
)
FLOAT = '<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="margin" w:horzAnchor="margin" w:tblpXSpec="left" w:tblpYSpec="bottom"/>'


def table(tag, width, floated=False):
    rows = ''.join(
        f'<w:tr><w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>'
        + para(f'{tag} row {i + 1}', ppr=SP, run_rpr=C)
        + '</w:tc></w:tr>'
        for i in range(3)
    )
    return (
        f'<w:tbl><w:tblPr>{FLOAT if floated else ""}<w:tblW w:w="{width}" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
        f'<w:tblBorders>{BORD}</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="{width}"/></w:tblGrid>{rows}</w:tbl>'
    )


SECT1 = f'<w:sectPr>{LETTER}</w:sectPr>'
body = (
    table('T1', 9360)
    + para('body one', ppr=SP, run_rpr=C)
    + table('F', 4000, floated=True)
    + para('anchor', ppr=SP + SECT1, run_rpr=C, mark_rpr='')
    + table('T2', 9360)
    + para('body two', ppr=SP, run_rpr=C)
)
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
