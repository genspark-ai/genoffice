import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Letter, Calibri 11 single, two nextPage sections. Section 1 is a cover: a title, three
# lines of text, a one-row tblpPr table floated to the margin bottom (tblpXSpec center,
# tblpYSpec bottom, tblOverlap never) and the empty paragraph carrying the sectPr right
# after it. Section 2 opens with an empty paragraph, a centred table and body text: the
# float is pinned at the bottom of page 1, the opener starts page 2 under it.
C = rpr('Calibri', 22)
T = rpr('Calibri', 48)
SP = spacing(after=0, line=240, rule='auto')
BORD = ''.join(
    f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')
)
FLOAT = (
    '<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="margin" w:horzAnchor="margin"'
    ' w:tblpXSpec="center" w:tblpYSpec="bottom"/><w:tblOverlap w:val="never"/>'
)


def table(tag, width, rows, floated=False, centered=False):
    body = ''.join(
        f'<w:tr><w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>'
        + para(f'{tag} row {i + 1}', ppr=SP, run_rpr=C)
        + '</w:tc></w:tr>'
        for i in range(rows)
    )
    jc = '<w:jc w:val="center"/>' if centered else ''
    return (
        f'<w:tbl><w:tblPr>{FLOAT if floated else ""}<w:tblW w:w="{width}" w:type="dxa"/>{jc}'
        f'<w:tblLayout w:type="fixed"/><w:tblBorders>{BORD}</w:tblBorders></w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="{width}"/></w:tblGrid>{body}</w:tbl>'
    )


SECT1 = f'<w:sectPr>{LETTER}</w:sectPr>'
body = (
    para('Cover title', ppr=SP, run_rpr=T)
    + para('cover line one', ppr=SP, run_rpr=C)
    + para('cover line two', ppr=SP, run_rpr=C)
    + para('cover line three', ppr=SP, run_rpr=C)
    + table('F', 6000, 1, floated=True)
    + para(None, ppr=SP + SECT1, mark_rpr='')
    + para(None, ppr=SP, mark_rpr=C)
    + table('T2', 6000, 3, centered=True)
    + para('body two', ppr=SP, run_rpr=C)
)
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
