import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# 11920 x 16860 page, 1133-twip margins: column 9654 twips (482.7 pt). TOC-style lines
# "title<tab>page" with a right stop 130 twips inside the edge (9524), on the edge (9654),
# past it (9798) and well inside (9510); a numbered hanging-indent entry like the corpus
# shape; and a title series whose end walks across the stop so Word's wrap threshold shows
SECT = '<w:pgSz w:w="11920" w:h="16860"/><w:pgMar w:top="1620" w:right="1133" w:bottom="280" w:left="1133" w:header="299" w:footer="491" w:gutter="0"/>'
SP = spacing(before=120, after=0, line=240, rule='auto')
TITLE = '<w:rPr><w:color w:val="1A1A1A"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>'
NUM = '<w:rPr><w:position w:val="4"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'


def tabs(right, leader='none'):
    return f'<w:tabs><w:tab w:val="left" w:pos="405" w:leader="none"/><w:tab w:val="right" w:pos="{right}" w:leader="{leader}"/></w:tabs>'


def entry(title, page, right=9524, leader='none', listed=False):
    num = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' if listed else ''
    ind = '<w:ind w:left="405" w:hanging="291"/>' if listed else '<w:ind w:left="405"/>'
    return (
        f'<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>{num}{tabs(right, leader)}{SP}{ind}<w:jc w:val="left"/></w:pPr>'
        f'<w:r>{TITLE}<w:t xml:space="preserve">{title}</w:t></w:r><w:r>{TITLE}<w:tab/></w:r>'
        f'<w:r>{NUM}<w:t>{page}</w:t></w:r></w:p>'
    )


LEVEL = lvl(fmt='decimalZero', text='%1', ind='<w:ind w:left="114" w:hanging="359"/>',
            rpr='<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:color w:val="8A2850"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>')
LIST_STYLE = '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="151"/><w:ind w:left="114"/></w:pPr></w:style>'

body = (
    entry('Quem conhece quem', '14', listed=True)
    + entry('Os 47 personagens', '7', listed=True)
    + entry('Figurino e objetos com memoria', '22', listed=True, leader='dot')
    + entry('Stop on the column edge', '14', right=9654)
    + entry('Stop past the column edge', '14', right=9798)
    + entry('Stop well inside the column', '14', right=9510)
    + entry('Leader to the near stop', '14', leader='dot')
    + para('Control body paragraph', ppr=SP + '<w:ind w:left="405"/>', run_rpr=TITLE)
)
# title end walks from ~455 pt to ~485 pt across the 9524 stop (476.2 pt) and the edge (482.7 pt)
for n in range(0, 15):
    body += entry('Near the stop ' + 'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM' + 'i' * n * 2, str(10 + n))
body += para('Series end', ppr=SP + '<w:ind w:left="405"/>', run_rpr=TITLE)
write(out_path(), body, styles_xml(doc_defaults('Arial', 22), LIST_STYLE), settings_xml(), sect=SECT,
      numbering=numbering_xml([[LEVEL]]))
