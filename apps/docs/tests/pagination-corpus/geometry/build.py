"""Synthetic geometry documents for the rendered pagination-invariant suite
(e2e/docs-pagination-invariants.spec.ts): floating tables (body, nested, header),
anchored pictures over every relH/relV x posOffset/align x wrap combination the
renderer branches on, an over-wide centred dxa table and a pct table with a
non-uniform grid. Letter, Calibri 11, single spacing, compat 15.

Usage: python3 build.py   (from this directory; rewrites the .docx files next to it)
"""
import os
import struct
import sys
import zipfile
import zlib

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '../../../../../packages/docx-engine/tests/fixtures/truths'))
from _docx import W, R, XML, doc_defaults, hdr_xml, para, rpr, settings_xml, spacing, styles_xml, write

C = rpr('Calibri', 22)
SP = spacing(after=0, line=240, rule='auto')
LOREM = (
    'The quick brown fox jumps over the lazy dog while the five boxing wizards jump quickly; '
    'pack my box with five dozen liquor jugs and sphinx of black quartz judge my vow. '
)
BORD = ''.join(
    f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')
)
DRAW_NS = (
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
)
IMG_RID = 'rId90'


def text(n=1, ppr=SP):
    return ''.join(para(LOREM * 3, ppr=ppr, run_rpr=C) for _ in range(n))


def cell(w, content):
    return f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>{content}</w:tc>'


def table(cols, rows, tblw='<w:tblW w:w="0" w:type="auto"/>', layout='<w:tblLayout w:type="fixed"/>', extra='', tcw=None):
    """cols: gridCol widths; rows: list of lists of cell paragraph XML"""
    tcw = tcw or cols
    grid = ''.join(f'<w:gridCol w:w="{c}"/>' for c in cols)
    trs = ''.join(
        '<w:tr>' + ''.join(cell(tcw[i], c) for i, c in enumerate(r)) + '</w:tr>' for r in rows
    )
    return (
        f'<w:tbl><w:tblPr>{extra}{tblw}{layout}<w:tblBorders>{BORD}</w:tblBorders></w:tblPr>'
        f'<w:tblGrid>{grid}</w:tblGrid>{trs}</w:tbl>'
    )


def tblp(vert='text', horz='margin', xspec=None, yspec=None, x=None, y=None, overlap=True):
    a = f' w:leftFromText="180" w:rightFromText="180" w:topFromText="0" w:bottomFromText="0" w:vertAnchor="{vert}" w:horzAnchor="{horz}"'
    a += f' w:tblpXSpec="{xspec}"' if xspec else f' w:tblpX="{x}"'
    a += f' w:tblpYSpec="{yspec}"' if yspec else f' w:tblpY="{y}"'
    return f'<w:tblpPr{a}/>' + ('' if overlap else '<w:tblOverlap w:val="never"/>')


def float_table(tag, width, pos, rows=3):
    return table([width], [[para(f'{tag} row {i + 1}', ppr=SP, run_rpr=C)] for i in range(rows)],
                 tblw=f'<w:tblW w:w="{width}" w:type="dxa"/>', extra=pos)


def picture(cx, cy, rel_h='column', rel_v='paragraph', x=0, y=0, align_h=None, align_v=None, wrap='square', pic_id=1):
    pos_h = f'<wp:align>{align_h}</wp:align>' if align_h else f'<wp:posOffset>{x}</wp:posOffset>'
    pos_v = f'<wp:align>{align_v}</wp:align>' if align_v else f'<wp:posOffset>{y}</wp:posOffset>'
    behind = 1 if wrap == 'behind' else 0
    wrap_xml = {
        'square': '<wp:wrapSquare wrapText="bothSides"/>',
        'topBottom': '<wp:wrapTopAndBottom/>',
        'behind': '<wp:wrapNone/>',
    }[wrap]
    return (
        f'<w:r><w:drawing {DRAW_NS}><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
        f'relativeHeight="{251658240 + pic_id}" behindDoc="{behind}" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="{rel_h}">{pos_h}</wp:positionH>'
        f'<wp:positionV relativeFrom="{rel_v}">{pos_v}</wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>{wrap_xml}'
        f'<wp:docPr id="{pic_id}" name="Picture {pic_id}"/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{pic_id}" name="Picture {pic_id}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{IMG_RID}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
        '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
    )


def anchor_para(pic_xml, lead='Anchor paragraph. ', ppr=SP):
    return f'<w:p><w:pPr>{ppr}</w:pPr>{pic_xml}<w:r>{C}<w:t xml:space="preserve">{lead}{LOREM * 3}</w:t></w:r></w:p>'


def png(w, h, rgb):
    raw = b''.join(b'\x00' + bytes(rgb) * w for _ in range(h))

    def chunk(tag, data):
        return struct.pack('>I', len(data)) + tag + data + struct.pack('>I', zlib.crc32(tag + data) & 0xFFFFFFFF)

    return (
        b'\x89PNG\r\n\x1a\n'
        + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0))
        + chunk(b'IDAT', zlib.compress(raw))
        + chunk(b'IEND', b'')
    )


def add_media(path, image):
    """append word/media/image1.png and its relationship to a written package"""
    src = zipfile.ZipFile(path)
    parts = {n: src.read(n) for n in src.namelist()}
    src.close()
    rels = parts['word/_rels/document.xml.rels'].decode()
    rel = f'<Relationship Id="{IMG_RID}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>'
    parts['word/_rels/document.xml.rels'] = rels.replace('</Relationships>', rel + '</Relationships>').encode()
    ct = parts['[Content_Types].xml'].decode()
    parts['[Content_Types].xml'] = ct.replace('<Override', '<Default Extension="png" ContentType="image/png"/><Override', 1).encode()
    parts['word/media/image1.png'] = image
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, data in parts.items():
            z.writestr(n, data)


STYLES = styles_xml(doc_defaults('Calibri', 22))
SETTINGS = settings_xml()


def emit(name, body, **kw):
    out = os.path.join(HERE, name)
    write(out, body, STYLES, SETTINGS, **kw)
    return out


# 1. body floating tables: text/margin anchored right, page/page absolute, margin/margin
#    with tblOverlap never, and one that lands near a page bottom (flow vs shift decision)
body = (
    text(1)
    + float_table('T1 text-right', 3200, tblp('text', 'margin', xspec='right', y=120))
    + text(3)
    + float_table('T2 page-abs', 3000, tblp('page', 'page', x=1000, y=9000))
    + text(3)
    + float_table('T3 margin-centre', 2600, tblp('margin', 'margin', xspec='center', y=2000, overlap=False))
    + text(2)
    + float_table('T4 inside', 2400, tblp('text', 'text', xspec='inside', yspec='top'))
    + text(4)
    + float_table('T5 bottom-of-page', 2800, tblp('margin', 'margin', xspec='right', y=11000), rows=6)
    + text(3)
)
emit('geo-float-tables-body.docx', body)

# 2. floating table nested in a cell + header/footer floating tables
nested_cell = (
    float_table('N1 nested', 3600, tblp('text', 'margin', xspec='left', y=0))
    + ''.join(para(LOREM, ppr=SP, run_rpr=C) for _ in range(4))
)
outer = table([9360], [[nested_cell], [text(1)]], tblw='<w:tblW w:w="9360" w:type="dxa"/>')
header = hdr_xml(
    float_table('H1 header', 2400, tblp('page', 'page', x=8000, y=300), rows=2)
    + para('Header text beside the floating table', ppr=SP, run_rpr=C),
    'hdr',
)
footer = hdr_xml(
    float_table('F1 footer', 2000, tblp('margin', 'margin', xspec='right', yspec='bottom'), rows=1)
    + para('Footer text', ppr=SP, run_rpr=C),
    'ftr',
)
LETTER_HF = (
    '<w:headerReference w:type="default" r:id="rIdheader1"/>'
    '<w:footerReference w:type="default" r:id="rIdfooter1"/>'
    '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
)
emit(
    'geo-float-tables-nested-hf.docx',
    text(2) + outer + text(5) + outer + text(2),
    sect=LETTER_HF,
    hf={'header1.xml': header, 'footer1.xml': footer},
)

# 3. anchored pictures: relH column/margin/page x posOffset/align, relV paragraph/page/margin,
#    wrap square/topBottom/behind; one anchor paragraph is indented (column origin vs indent)
EMU = 12700
pics = [
    dict(rel_h='column', x=0, y=0, wrap='square'),
    dict(rel_h='column', align_h='right', y=0, wrap='square'),
    dict(rel_h='margin', x=200 * EMU, y=0, wrap='topBottom'),
    dict(rel_h='margin', align_h='center', y=0, wrap='topBottom'),
    dict(rel_h='page', x=60 * EMU, y=0, wrap='behind'),
    dict(rel_h='page', align_h='right', y=0, wrap='square'),
    dict(rel_h='column', x=0, rel_v='page', y=600 * EMU, wrap='square'),
    dict(rel_h='margin', align_h='left', rel_v='margin', align_v='bottom', wrap='topBottom'),
    dict(rel_h='page', align_h='center', rel_v='margin', y=200 * EMU, wrap='behind'),
    dict(rel_h='column', align_h='center', rel_v='paragraph', y=0, wrap='topBottom'),
]
IND = SP + '<w:ind w:left="1276" w:hanging="0"/>'
body = text(1)
for i, p in enumerate(pics):
    ppr = IND if i in (0, 4, 9) else SP
    body += anchor_para(picture(144 * EMU, 96 * EMU, pic_id=i + 1, **p), lead=f'P{i + 1} ', ppr=ppr) + text(2)
out = emit('geo-anchored-pictures.docx', body)
add_media(out, png(144, 96, (40, 90, 180)))

# 4. over-wide centred dxa table (12704 twips on a 9360 column) with a nested dxa table
nested = table([5000], [[para('nested 5000 dxa', ppr=SP, run_rpr=C)]], tblw='<w:tblW w:w="5000" w:type="dxa"/>')
wide_rows = [[para(f'c{r}{c}', ppr=SP, run_rpr=C) for c in range(3)] for r in range(4)]
wide_rows[1][1] = nested + para('after nested', ppr=SP, run_rpr=C)
wide = table(
    [4234, 4235, 4235],
    wide_rows,
    tblw='<w:tblW w:w="12704" w:type="dxa"/>',
    extra='<w:jc w:val="center"/>',
)
emit('geo-table-overwide-centred.docx', text(2) + wide + text(3) + wide + text(2))

# 5. pct tables with a non-uniform Word-saved grid: 100% (grid sums to the column)
#    and 60% (grid sums to 60% of it), no tblLayout
pct_rows = [[para(f'r{r}c{c} ' + LOREM[: 40 * (c + 1)], ppr=SP, run_rpr=C) for c in range(3)] for r in range(5)]
pct100 = table([1200, 3000, 5160], pct_rows, tblw='<w:tblW w:w="5000" w:type="pct"/>', layout='')
pct60 = table([900, 1716, 3000], pct_rows, tblw='<w:tblW w:w="3000" w:type="pct"/>', layout='')
emit('geo-table-pct-nonuniform.docx', text(1) + pct100 + text(2) + pct60 + text(3) + pct100 + text(1))
