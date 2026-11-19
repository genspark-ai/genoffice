import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Anchored picture whose posOffset X + extent runs past the column's right edge
# (into the right margin or off the paper): does Word keep the declared offset
# or pull the picture back into the column? One case per A4 page, a labelled
# anchor paragraph followed by body text; the pictures are red bitmaps.
WP = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
PIC = 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
PT = 12700
CX, CY = 207 * PT, 129 * PT
BODY = ('Uses heat and pressure to activate a thermoplastic bonding layer and join fabric '
        'components so the seam stays flat and light. ') * 3


def anchor(i, rel_h, x, wrap, cx=CX, cy=CY, wrap_text='bothSides'):
    pos = lambda v: f'<wp:posOffset>{v}</wp:posOffset>' if isinstance(v, int) else f'<wp:align>{v}</wp:align>'
    wrap_el = ('<wp:wrapNone/>' if wrap == 'None' else '<wp:wrapTopAndBottom/>' if wrap == 'TopAndBottom'
               else f'<wp:wrap{wrap} wrapText="{wrap_text}"/>')
    return (
        f'<w:r><w:drawing><wp:anchor {WP} distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="{251658240 + i}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="{rel_h}">{pos(x)}</wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>82850</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>{wrap_el}<wp:docPr id="{i + 1}" name="pic{i}"/>'
        f'<a:graphic {A}><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic {PIC}><pic:nvPicPr><pic:cNvPr id="0" name="pic{i}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="rIdImg0"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        '</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
    )


def run(t, bold=False):
    return f'<w:r>{"<w:rPr><w:b/></w:rPr>" if bold else ""}<w:t xml:space="preserve">{t}</w:t></w:r>'


SP = spacing(after=160, line=259, rule='auto')
PB = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
cases = [
    ('C1 margin 293pt square', 'margin', 293 * PT, 'Square', {}),
    ('C2 column 293pt square', 'column', 293 * PT, 'Square', {}),
    ('C3 margin 400pt square off-paper', 'margin', 400 * PT, 'Square', {}),
    ('C4 margin 293pt square wrapText left', 'margin', 293 * PT, 'Square', {'wrap_text': 'left'}),
    ('C5 margin 293pt topBottom', 'margin', 293 * PT, 'TopAndBottom', {}),
    ('C6 margin 200pt square control', 'margin', 200 * PT, 'Square', {}),
    ('C7 column -40pt square left overhang', 'column', -40 * PT, 'Square', {}),
    ('C8 margin 0 square 500pt wide', 'margin', 0, 'Square', {'cx': 500 * PT, 'cy': 100 * PT}),
    ('C9 margin 293pt wrapNone', 'margin', 293 * PT, 'None', {}),
    ('C10 margin 293pt square 60pt tall', 'margin', 293 * PT, 'Square', {'cy': 60 * PT}),
]
body = ''
for i, (label, rel_h, x, wrap, kw) in enumerate(cases):
    if i:
        body += PB
    body += f'<w:p><w:pPr>{SP}</w:pPr>{anchor(i, rel_h, x, wrap, **kw)}{run(label, True)}</w:p>'
    for _ in range(4):
        body += f'<w:p><w:pPr>{SP}</w:pPr>{run(BODY)}</w:p>'
# two side-wrapped pictures of consecutive anchor paragraphs whose boxes overlap
# vertically (the second anchor paragraph starts above the first picture's bottom)
SHORT = 'Uses heat and pressure to activate a thermoplastic layer. '
pairs = [
    ('D1 right 293pt then right 293pt', 'margin', 293 * PT, 'margin', 293 * PT),
    ('D2 right 200pt then right 200pt in column', 'margin', 200 * PT, 'margin', 200 * PT),
    ('D3 left 0 then right 293pt', 'column', 0, 'margin', 293 * PT),
]
for j, (label, rel_a, xa, rel_b, xb) in enumerate(pairs):
    n = len(cases) + 2 * j
    body += PB
    body += f'<w:p><w:pPr>{SP}</w:pPr>{anchor(n, rel_a, xa, "Square")}{run(label + " first", True)}</w:p>'
    for _ in range(3):
        body += f'<w:p><w:pPr>{SP}</w:pPr>{run(SHORT)}</w:p>'
    body += f'<w:p><w:pPr>{SP}</w:pPr>{anchor(n + 1, rel_b, xb, "Square")}{run(label + " second", True)}</w:p>'
    for _ in range(4):
        body += f'<w:p><w:pPr>{SP}</w:pPr>{run(BODY)}</w:p>'
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml(), sect=A4, media=[PNG_1X1])
