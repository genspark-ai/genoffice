import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Legacy layout (compatibilityMode 14, same as no setting): a wrapSquare picture anchored
# to the last paragraph that fits a page stays on that page and overflows the bottom
# margin (pulled up only at the paper edge); only a column-wide picture whose anchor
# line no longer fits below it turns the page. One case per nextPage section: A4,
# pgMar top 509 / bottom 851 (body 25.45 .. 799.37 pt), Arial 11 line 240 after 0,
# 39 filler lines put the anchor paragraph top at 518.8 pt (280.5 pt of body left).
WP = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
PIC = 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
PT = 12700
A4_509 = ('<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="509" w:right="851" w:bottom="851" '
          'w:left="851" w:header="0" w:footer="387" w:gutter="0"/>')
C = rpr('Arial', 22)
SP = spacing(after=0, line=240, rule='auto')


def anchor(i, x, height, width):
    cx, cy = int(width * PT), int(height * PT)
    return (
        f'<w:r><w:drawing><wp:anchor {WP} distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="{251658240 + i}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="column"><wp:posOffset>{x}</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>139065</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="{i + 1}" name="pic{i}"/>'
        f'<a:graphic {A}><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic {PIC}><pic:nvPicPr><pic:cNvPr id="0" name="pic{i}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="rIdImg0"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        '</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
    )


NARROW, WIDE = 1688465, -288925  # column offsets: 133 pt in, 22.75 pt into the left margin
cases = [
    ('A', NARROW, 299.5, 299.5, ''),  # ends 829.2 > body 799.4: stays, overflows the margin
    ('B', NARROW, 299.5, 299.5, 'Anchor text'),  # same with text in the anchor paragraph
    ('C', NARROW, 330, 299.5, ''),  # would end 859.7 > paper 841.9: pulled up to the paper edge
    ('E', WIDE, 200, 557.5, ''),  # column-wide, ends 729.7: the anchor line lands below it
    ('D', WIDE, 330, 557.5, ''),  # column-wide, its anchor line cannot land: turns the page
]
body = ''
for i, (tag, x, height, width, text) in enumerate(cases):
    for n in range(39):
        body += para(f'{tag} filler line {n + 1}', ppr=SP, run_rpr=C)
    run = f'<w:r>{C}<w:t>{text}</w:t></w:r>' if text else ''
    body += f'<w:p><w:pPr>{SP}{C}</w:pPr>{anchor(i, x, height, width)}{run}</w:p>'
    body += para(f'{tag} NEXT paragraph', ppr=SP, run_rpr=C)
    last = i == len(cases) - 1
    sect = '' if last else f'<w:sectPr><w:type w:val="nextPage"/>{A4_509}</w:sectPr>'
    body += para(f'{tag} MARK', ppr=SP, run_rpr=C, mark_rpr=C + sect)
write(out_path(), body, styles_xml(doc_defaults('Arial', 22)), settings_xml(compat_mode=14), sect=A4_509, media=[PNG_1X1])
