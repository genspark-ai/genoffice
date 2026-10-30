import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Letter, docDefaults Aptos 12, pPrDefault after 160 line 276 auto. Three outer one-cell
# tables, each holding a nested table: (A) seven rows whose paragraphs carry no spacing
# of their own and inherit the default after; (B) the same rows with direct after 0;
# (C) one row whose six-line Calibri 11 paragraph takes a style with line 254 auto, a
# multiple the nested cell must keep.
A12 = rpr('Aptos', 24)
C11 = rpr('Calibri', 22)
DEF = spacing(after=160, line=276, rule='auto')
BORD = ''.join(
    f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')
)
STYLE = (
    '<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/><w:basedOn w:val="Normal"/>'
    '<w:pPr><w:spacing w:after="0" w:line="254" w:lineRule="auto"/></w:pPr></w:style>'
)


def tbl(width, rows, extra=''):
    return (
        f'<w:tbl><w:tblPr>{extra}<w:tblW w:w="{width}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders>{BORD}</w:tblBorders></w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="{width}"/></w:tblGrid>' + rows + '</w:tbl>'
    )


def cell(width, content):
    return f'<w:tr><w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>{content}</w:tc></w:tr>'


def outer(inner):
    return tbl(9360, cell(9360, inner + para(None, run_rpr=A12)))


nested_a = tbl(4000, ''.join(cell(4000, para(f'row {i + 1}', run_rpr=A12)) for i in range(7)))
nested_b = tbl(4000, ''.join(cell(4000, para(f'row {i + 1}', ppr=spacing(after=0), run_rpr=A12)) for i in range(7)))
nested_c = tbl(4000, cell(4000, para([f'line {i + 1}' for i in range(6)], ppr='<w:pStyle w:val="Body"/>', run_rpr=C11)))
body = (
    para('lead', run_rpr=A12)
    + outer(nested_a)
    + para('mid', run_rpr=A12)
    + outer(nested_b)
    + para('mid two', run_rpr=A12)
    + outer(nested_c)
    + para('tail', run_rpr=A12)
)
write(out_path(), body, styles_xml(doc_defaults('Aptos', 24, ppr=DEF), STYLE), settings_xml())
