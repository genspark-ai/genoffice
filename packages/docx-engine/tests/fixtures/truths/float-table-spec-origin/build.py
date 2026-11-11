import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Floating-table placement: (A) w:tblpXSpec keywords against the margin / page /
# text bands (asymmetric margins so the two bands differ), (B) the anchor
# paragraph beside a text-anchored float when the side room is narrower than
# its next word, (D) the same keywords on a table nested in a cell and on
# header tables. One case per page; tables are one fixed cell, Calibri 11.
SECT = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="720" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
PB = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
LONG = 'Internationalization'
WORDS = 'anchor paragraph text that keeps going with enough plain words to wrap beside a floating table for several lines on end ' * 3


def ftbl(tblp, width=3000, label='T', rows=1):
    return tbl([para(label)] * rows, width=width, extra=f'<w:tblpPr {tblp}/>')


def case(label, tblp, width=3000, anchor_text=None, ppr='', rows=1):
    txt = anchor_text if anchor_text is not None else WORDS
    return PB + para(label) + ftbl(tblp, width, rows=rows) + para(txt, ppr) + para('after ' + label)


body = para('F0 lead')
# (A) keyword X on the margin / page / text bands, text-anchored Y
a_cases = [
    ('A1 margin right', 'w:horzAnchor="margin" w:tblpXSpec="right" w:vertAnchor="text" w:tblpY="1"'),
    ('A2 page right', 'w:horzAnchor="page" w:tblpXSpec="right" w:vertAnchor="text" w:tblpY="1"'),
    ('A3 margin center', 'w:horzAnchor="margin" w:tblpXSpec="center" w:vertAnchor="text" w:tblpY="1"'),
    ('A4 page center', 'w:horzAnchor="page" w:tblpXSpec="center" w:vertAnchor="text" w:tblpY="1"'),
    ('A5 margin inside', 'w:horzAnchor="margin" w:tblpXSpec="inside" w:vertAnchor="text" w:tblpY="1"'),
    ('A6 margin outside', 'w:horzAnchor="margin" w:tblpXSpec="outside" w:vertAnchor="text" w:tblpY="1"'),
    ('A8 page left', 'w:horzAnchor="page" w:tblpXSpec="left" w:vertAnchor="text" w:tblpY="1"'),
    ('A10 text right numeric', 'w:horzAnchor="text" w:tblpX="7000" w:vertAnchor="text" w:tblpY="1"'),
]
for label, tblp in a_cases:
    body += case(label, tblp)
body += case('A7 text right indented anchor', 'w:horzAnchor="text" w:tblpXSpec="right" w:vertAnchor="text" w:tblpY="1"', ppr='<w:ind w:left="720"/>')
body += case('A9 margin right overwide', 'w:horzAnchor="margin" w:tblpXSpec="right" w:vertAnchor="text" w:tblpY="1"', width=11000, rows=3)
# (B) side room vs the anchor paragraph's next word; left float at the column edge,
# rightFromText 180 (9 pt); column 504 pt
b_cases = [
    ('B1 room 30 long word', 9120, LONG + ' ' + WORDS),
    ('B2 room 120 long word', 7320, LONG + ' ' + WORDS),
    ('B3 room 30 short then long', 9120, 'a ' + LONG + ' ' + WORDS),
    ('B4 room 60 word 50', 8520, 'Mountains ' + WORDS),
    ('B5 room 30 words only', 9120, WORDS),
]
for label, w, txt in b_cases:
    body += case(label, 'w:horzAnchor="margin" w:tblpX="0" w:vertAnchor="text" w:tblpY="1" w:rightFromText="180"', width=w, anchor_text=txt, rows=3)
# (D) nested: a one-cell outer table spanning the column, keyword floats inside
inner = lambda tblp, w=3000: tbl([para('N')], width=w, extra=f'<w:tblpPr {tblp}/>')
for label, tblp in [
    ('D1 nested margin right', 'w:horzAnchor="margin" w:tblpXSpec="right" w:vertAnchor="text" w:tblpY="1"'),
    ('D2 nested margin center', 'w:horzAnchor="margin" w:tblpXSpec="center" w:vertAnchor="text" w:tblpY="1"'),
    ('D3 nested page right', 'w:horzAnchor="page" w:tblpXSpec="right" w:vertAnchor="text" w:tblpY="1"'),
]:
    cell = inner(tblp) + para(WORDS) + para('cell tail')
    body += PB + para(label) + tbl([cell], width=10080) + para('after ' + label)
body += para('F end')

# headers: section per header case (nextPage sections carry their own header)
hdr_cases = [
    ('header1.xml', 'w:horzAnchor="margin" w:tblpXSpec="right" w:vertAnchor="text" w:tblpY="1"', 3000),
    ('header2.xml', 'w:horzAnchor="page" w:tblpXSpec="center" w:vertAnchor="text" w:tblpY="1"', 3000),
    ('header3.xml', 'w:horzAnchor="margin" w:tblpX="0" w:vertAnchor="margin" w:tblpY="0"', 3000),
]
hf = {}
secs = ''
for i, (name, tblp, w) in enumerate(hdr_cases):
    hf[name] = hdr_xml(ftbl(tblp, w, label=f'H{i + 1}') + para(f'header line {i + 1}'))
    sect_pr = f'<w:sectPr><w:headerReference w:type="default" r:id="rId{name[:-4]}"/><w:type w:val="nextPage"/>{SECT}</w:sectPr>'
    secs += f'<w:p><w:pPr>{sect_pr}</w:pPr>' + f'<w:r><w:t>H{i + 1} body line</w:t></w:r></w:p>'
body = body + secs
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml(), sect=SECT, hf=hf)
