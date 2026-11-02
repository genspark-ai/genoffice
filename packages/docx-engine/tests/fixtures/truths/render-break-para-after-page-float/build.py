import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# A4, Calibri 11, after 160 / line 259 (Word 2013+ Normal). Cover page: a 28pt title,
# empty lines, a six-row page-anchored tblpPr table (vertAnchor page, tblpY 7500 =
# 375pt; its band 375-620pt reaches past the paragraphs that follow it in the flow),
# then the empty keepNext/keepLines paragraph that anchors it and a paragraph holding
# only a page break. Page 2: a keepNext heading, an empty paragraph, an 8-row table,
# another break-only paragraph; page 3: a second heading and body text.
A4 = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1134" w:bottom="1134" w:left="1134" w:header="284" w:footer="708" w:gutter="0"/>'
C = rpr('Calibri', 22)
H = rpr('Calibri', 32)
T = rpr('Calibri', 56, extra='<w:b/>')
SP = spacing(after=0, line=240, rule='auto')
HEAD = '<w:keepNext/><w:keepLines/>' + spacing(before=240, after=0)
BORD = ''.join(
    f'<w:{s} w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
    for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')
)
FLOAT = (
    '<w:tblpPr w:leftFromText="142" w:rightFromText="142" w:bottomFromText="198"'
    ' w:vertAnchor="page" w:horzAnchor="margin" w:tblpY="7500"/>'
)
COLS = [1686, 993, 2551, 1701, 2552]
LABELS = ['Date', 'Version', 'Compiled', 'Verified', 'Approved']
PB = para(None, ppr='', mark_rpr=H).replace('</w:pPr>', '</w:pPr><w:r><w:br w:type="page"/></w:r>')


def sig_row(texts):
    cells = ''.join(
        f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>'
        + para(t, ppr=SP + '<w:jc w:val="center"/>', run_rpr=C)
        + '</w:tc>'
        for w, t in zip(COLS, texts)
    )
    return f'<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="533"/></w:trPr>{cells}</w:tr>'


def sig_table():
    grid = ''.join(f'<w:gridCol w:w="{w}"/>' for w in COLS)
    rows = ''.join(
        sig_row([f'0{i}-01-2026', str(6 - i), 'Author Name', ['Reviewer Name', '(Engineer)'], ['Approver Name', '(Manager)']])
        for i in range(1, 6)
    ) + sig_row(LABELS)
    return (
        f'<w:tbl><w:tblPr>{FLOAT}<w:tblW w:w="{sum(COLS)}" w:type="dxa"/><w:tblInd w:w="0" w:type="dxa"/>'
        f'<w:tblBorders>{BORD}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>'
        f'<w:tblGrid>{grid}</w:tblGrid>{rows}</w:tbl>'
    )


def rev_table(rows):
    widths = [988, 7096, 1544]
    grid = ''.join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    trs = ''
    for i in range(rows):
        texts = [str(i), f'Revision {i} description text for the change log entry', '01-01-2026']
        trs += '<w:tr>' + ''.join(
            f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>'
            + para(t, ppr=spacing(before=60, after=60), run_rpr=C)
            + '</w:tc>'
            for w, t in zip(widths, texts)
        ) + '</w:tr>'
    return (
        f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>{BORD}</w:tblBorders></w:tblPr>'
        f'<w:tblGrid>{grid}</w:tblGrid>{trs}</w:tbl>'
    )


body = (
    para(None, mark_rpr=T) * 2
    + para('Cover Title', ppr=SP + '<w:jc w:val="center"/>', run_rpr=T)
    + para('Subtitle line', ppr=SP + '<w:jc w:val="center"/>', run_rpr=T)
    + para(None, mark_rpr=C) * 2
    + sig_table()
    + para(None, ppr=HEAD, mark_rpr=H)
    + PB
    + para('Revision history', ppr=HEAD, run_rpr=H)
    + para(None, mark_rpr=C)
    + rev_table(8)
    + PB
    + para('Contents', ppr=HEAD, run_rpr=H)
    + para('Body text after the contents heading.', run_rpr=C)
)
write(
    out_path(),
    body,
    styles_xml(doc_defaults('Calibri', 22, ppr=spacing(after=160, line=259, rule='auto'))),
    settings_xml(),
    sect=A4,
)
