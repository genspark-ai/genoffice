import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Anchored-picture origins: (1) relativeFrom="column" / "character" X in an
# indented paragraph (does w:ind move the picture?), (2) a margin-relative
# posOffset <= 0 under empty paragraphs (does the picture lift over them?),
# (3) wp:align keywords (left/center/right x top/center/bottom) against the
# margin and page bands for wrapSquare and wrapTopAndBottom. One case per page
# for (2) and (3); the pictures are 1 x 0.5 in red bitmaps.
WP = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
PIC = 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
CX, CY = 914400, 457200


def anchor(i, rel_h='column', rel_v='paragraph', x=0, y=0, wrap='None', behind=0):
    pos = lambda v: f'<wp:posOffset>{v}</wp:posOffset>' if isinstance(v, int) else f'<wp:align>{v}</wp:align>'
    wrap_el = '<wp:wrapNone/>' if wrap == 'None' else '<wp:wrapTopAndBottom/>' if wrap == 'TopAndBottom' else f'<wp:wrap{wrap} wrapText="bothSides"/>'
    return (
        f'<w:r><w:drawing><wp:anchor {WP} distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{251658240 + i}" behindDoc="{behind}" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="{rel_h}">{pos(x)}</wp:positionH><wp:positionV relativeFrom="{rel_v}">{pos(y)}</wp:positionV>'
        f'<wp:extent cx="{CX}" cy="{CY}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>{wrap_el}<wp:docPr id="{i + 1}" name="pic{i}"/>'
        f'<a:graphic {A}><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic {PIC}><pic:nvPicPr><pic:cNvPr id="0" name="pic{i}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="rIdImg0"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{CX}" cy="{CY}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        '</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
    )


def run(t):
    return f'<w:r><w:t xml:space="preserve">{t}</w:t></w:r>'


SEP = spacing(after=1000)
IND = '<w:ind w:left="720" w:firstLine="360"/>'
HANG = '<w:ind w:left="720" w:hanging="360"/>'
n = 0
body = ''
# (1) column / character X under an indent, one anchor paragraph each
cases1 = [
    ('X1a column wrapNone ind', IND, 'column', 914400, 'None'),
    ('X1b column wrapSquare ind', IND, 'column', 914400, 'Square'),
    ('X1c column wrapSquare hanging', HANG, 'column', 914400, 'Square'),
    ('X1d character wrapNone ind', IND, 'character', 914400, 'None'),
    ('X1e character wrapSquare no ind', '', 'character', 914400, 'Square'),
    ('X1f column wrapNone no ind control', '', 'column', 457200, 'None'),
    ('X1g margin wrapNone ind', IND, 'margin', 914400, 'None'),
]
for label, ind, rel_h, x, wrap in cases1:
    body += f'<w:p><w:pPr>{SEP}{ind}</w:pPr>{run(label + " ")}{anchor(n, rel_h, "paragraph", x, 0, wrap)}{run(" tail")}</w:p>'
    n += 1
# picture-only paragraph under an indent
body += f'<w:p><w:pPr>{SEP}{IND}</w:pPr>{anchor(n, "column", "paragraph", 914400, 0, "Square")}</w:p>'
n += 1
body += para('X1 end')
PB = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
# (2) margin-relative posOffset under three empty paragraphs
for label, y, wrap in [('L1 margin 0 topBottom', 0, 'TopAndBottom'), ('L2 margin -18pt topBottom', -228600, 'TopAndBottom'), ('L3 margin 0 square', 0, 'Square'), ('L4 margin 20pt topBottom', 254000, 'TopAndBottom')]:
    body += PB + '<w:p/><w:p/><w:p/>'
    body += f'<w:p>{run(label + " anchor")}{anchor(n, "column", "margin", 0, y, wrap)}</w:p>'
    body += para('after ' + label)
    n += 1
# (3) wp:align keywords
cases3 = []
for wrap in ('Square', 'TopAndBottom'):
    for ah in ('left', 'center', 'right'):
        for av in ('top', 'center', 'bottom'):
            cases3.append(('margin', 'margin', ah, av, wrap))
    cases3 += [('page', 'page', 'center', 'top', wrap), ('page', 'page', 'right', 'bottom', wrap), ('page', 'page', 'center', 'center', wrap), ('margin', 'page', 'center', 'center', wrap), ('column', 'margin', 'center', 'top', wrap), ('column', 'paragraph', 'center', 'top', wrap)]
for rel_h, rel_v, ah, av, wrap in cases3:
    label = f'A {rel_h}/{rel_v} {ah}/{av} {wrap}'
    body += PB + para('line 1') + para('line 2')
    body += f'<w:p>{run(label + " anchor")}{anchor(n, rel_h, rel_v, ah, av, wrap)}</w:p>'
    body += para('after ' + label)
    n += 1
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml(), media=[PNG_1X1])
