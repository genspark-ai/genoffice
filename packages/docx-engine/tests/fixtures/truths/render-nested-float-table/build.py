import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Letter, Calibri 11 single. Outer 1-column table, two rows of trHeight atLeast 3000 with
# tblCellMar 108 all round; each cell holds a page-anchored tblpPr table (tblpY 261,
# horzAnchor margin, tblOverlap never, 4000 wide) followed by five empty paragraphs
# that wrap beside the float instead of stacking under it.
C = rpr('Calibri', 22)
SP = spacing(after=0, line=240, rule='auto')
BORD = ''.join(
    f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')
)


def nested(tag):
    rows = ''.join(
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
        + para(f'{tag} nested {i + 1}', ppr=SP, run_rpr=C)
        + '</w:tc></w:tr>'
        for i in range(2)
    )
    return (
        '<w:tbl><w:tblPr>'
        '<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="page" w:horzAnchor="margin" w:tblpXSpec="left" w:tblpY="261"/>'
        '<w:tblOverlap w:val="never"/><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
        f'<w:tblBorders>{BORD}</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>{rows}</w:tbl>'
    )


def outer_row(tag):
    cell = nested(tag) + ''.join(para(None, ppr=SP, run_rpr=C) for _ in range(5))
    return (
        '<w:tr><w:trPr><w:trHeight w:val="3000" w:hRule="atLeast"/></w:trPr>'
        '<w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>' + cell + '</w:tc></w:tr>'
    )


MAR = ''.join(f'<w:{s} w:w="108" w:type="dxa"/>' for s in ('top', 'left', 'bottom', 'right'))
outer = (
    f'<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
    f'<w:tblBorders>{BORD}</w:tblBorders><w:tblCellMar>{MAR}</w:tblCellMar></w:tblPr>'
    '<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>' + outer_row('A') + outer_row('B') + '</w:tbl>'
)
body = para('lead', ppr=SP, run_rpr=C) + outer + para('tail', ppr=SP, run_rpr=C)
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
