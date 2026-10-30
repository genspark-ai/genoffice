import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import zipfile
from _docx import *
from _hfprobe import hf_doc, run, hp
# Header paragraphs around a side-wrapped drawing: a page-anchored wrapSquare
# picture in the first header paragraph, one nextPage section per case, one
# Calibri body line per section so each page's body top measures its strip.
# CASE00 control (no drawing); CASE01 a 540 pt wide rule across the whole
# column at page y 80 pt (header line 4 cannot sit beside it); CASE02 a 100 pt
# wide rule at the column's left edge, same y (the line flows beside it);
# CASE03 the full-width rule anchored paragraph-relative at the same page y;
# CASE04 a 100 pt rule in the middle of the column (gaps on both sides).
WP = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
PIC = 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
PT = 12700
C = rpr('Calibri', 22)
SP = spacing(after=0, line=240, rule='auto')


def anchor(i, rel_v, x_pt, y_pt, w_pt, h_pt=1.5):
    return (
        f'<w:r><w:drawing><wp:anchor {WP} distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{251658240 + i}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="page"><wp:posOffset>{int(x_pt * PT)}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="{rel_v}"><wp:posOffset>{int(y_pt * PT)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{int(w_pt * PT)}" cy="{int(h_pt * PT)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:wrapSquare wrapText="bothSides"/><wp:docPr id="{i + 1}" name="rule{i}"/>'
        f'<a:graphic {A}><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic {PIC}><pic:nvPicPr><pic:cNvPr id="0" name="rule{i}"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rIdImg0"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(w_pt * PT)}" cy="{int(h_pt * PT)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        '</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
    )


LINES = ['UNIVERSITY HEADER LINE ONE', 'Faculty header line two', 'Department header line three', 'Programme header line four']


def header(drawing=''):
    out = ''
    for k, text in enumerate(LINES):
        out += hp(SP, (drawing if k == 0 else '') + run(text, C))
    return out


cases = [
    header(),
    header(anchor(1, 'page', 36, 80, 540)),
    header(anchor(2, 'page', 36, 80, 100)),
    header(anchor(3, 'paragraph', 36, 44, 540)),
    header(anchor(4, 'page', 236, 80, 100)),
]
out = out_path()
hf_doc(out, cases)
# header parts carry their own rels: point every header at the shared picture
rel = (
    f'{XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rIdImg0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image0.png"/>'
    '</Relationships>'
)
with zipfile.ZipFile(out, 'a', zipfile.ZIP_DEFLATED) as z:
    names = set(z.namelist())
    for i in range(1, len(cases) + 1):
        z.writestr(f'word/_rels/header{i}.xml.rels', rel)
    if 'word/media/image0.png' not in names:
        z.writestr('word/media/image0.png', PNG_1X1)
    if '[Content_Types].xml' in names:
        ct = z.read('[Content_Types].xml').decode()
# zipfile cannot replace entries in place: rewrite the package with the png
# default and the r: namespace the header anchors' r:embed needs
tmp = out + '.tmp'
with zipfile.ZipFile(out) as src, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as dst:
    for item in src.infolist():
        data = src.read(item.filename)
        if item.filename == '[Content_Types].xml' and '<Default Extension="png"' not in ct:
            data = data.decode().replace('<Default Extension="xml"', '<Default Extension="png" ContentType="image/png"/><Default Extension="xml"').encode()
        if item.filename.startswith('word/header'):
            data = data.decode().replace(f'<w:hdr {W}>', f'<w:hdr {W} {R}>').encode()
        dst.writestr(item, data)
os.replace(tmp, out)
