"""Minimal synthetic docx writer shared by the truth fixtures' build.py scripts.

Usage: python3 build.py <out.docx>  (from the fixture directory)
"""
import base64
import sys
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
LETTER = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
A4 = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>'


def rpr(font=None, sz=None, east_asia=None, extra=''):
    fonts = ''
    if font or east_asia:
        attrs = ''
        if font:
            attrs += f' w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"'
        if east_asia:
            attrs += f' w:eastAsia="{east_asia}"'
        fonts = f'<w:rFonts{attrs}/>'
    size = f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>' if sz else ''
    body = fonts + size + extra
    return f'<w:rPr>{body}</w:rPr>' if body else ''


def para(text, ppr='', run_rpr='', mark_rpr=None):
    mark = mark_rpr if mark_rpr is not None else run_rpr
    pp = f'<w:pPr>{ppr}{mark}</w:pPr>' if (ppr or mark) else ''
    runs = ''
    if text is not None:
        lines = text if isinstance(text, list) else [text]
        parts = []
        for i, line in enumerate(lines):
            if i:
                parts.append(f'<w:r>{run_rpr}<w:br/></w:r>')
            parts.append(f'<w:r>{run_rpr}<w:t xml:space="preserve">{line}</w:t></w:r>')
        runs = ''.join(parts)
    return f'<w:p>{pp}{runs}</w:p>'


def spacing(before=None, after=None, line=None, rule=None, extra=''):
    a = ''
    if before is not None:
        a += f' w:before="{before}"'
    if after is not None:
        a += f' w:after="{after}"'
    if line is not None:
        a += f' w:line="{line}"'
    if rule:
        a += f' w:lineRule="{rule}"'
    return f'<w:spacing{a}{extra}/>'


def doc_defaults(font='Calibri', sz=22, east_asia=None, ppr=spacing(after=0, line=240, rule='auto')):
    return (
        f'<w:docDefaults><w:rPrDefault>{rpr(font, sz, east_asia)}</w:rPrDefault>'
        f'<w:pPrDefault><w:pPr>{ppr}</w:pPr></w:pPrDefault></w:docDefaults>'
    )


def styles_xml(defaults, extra_styles=''):
    return (
        f'{XML}<w:styles {W}>{defaults}'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        f'{extra_styles}</w:styles>'
    )


def settings_xml(compat_flags='', compat_mode=15):
    return (
        f'{XML}<w:settings {W}><w:compat>{compat_flags}'
        f'<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat_mode}"/>'
        '</w:compat></w:settings>'
    )


def font_table_xml(fonts):
    """fonts: list of (name, dict of child tag -> val) e.g. ('Mont', {'panose1': '0203...', 'charset': '00', 'family': 'roman'})"""
    out = ''
    for name, props in fonts:
        kids = ''.join(f'<w:{k} w:val="{v}"/>' for k, v in props.items())
        out += f'<w:font w:name="{name}">{kids}</w:font>'
    return f'{XML}<w:fonts {W}>{out}</w:fonts>'


def numbering_xml(levels_by_abstract):
    """levels_by_abstract: list of lists of level XML strings (w:lvl bodies); numId i+1 maps to abstractNum i"""
    out = ''
    for i, levels in enumerate(levels_by_abstract):
        lvls = ''.join(f'<w:lvl w:ilvl="{k}">{l}</w:lvl>' for k, l in enumerate(levels))
        out += f'<w:abstractNum w:abstractNumId="{i}"><w:multiLevelType w:val="hybridMultilevel"/>{lvls}</w:abstractNum>'
    for i in range(len(levels_by_abstract)):
        out += f'<w:num w:numId="{i + 1}"><w:abstractNumId w:val="{i}"/></w:num>'
    return f'{XML}<w:numbering {W}>{out}</w:numbering>'


def lvl(fmt='decimal', text='%1.', ind='', rpr='', extra=''):
    return f'<w:start w:val="1"/><w:numFmt w:val="{fmt}"/><w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/><w:pPr>{extra}{ind}</w:pPr>{rpr}'


def theme_xml(major='Calibri Light', minor='Calibri', major_ea='', minor_ea='', major_cs='', minor_cs=''):
    A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    def fonts(tag, lat, ea, cs):
        return f'<a:{tag}><a:latin typeface="{lat}"/><a:ea typeface="{ea}"/><a:cs typeface="{cs}"/></a:{tag}>'
    return (
        f'{XML}<a:theme {A} name="Office"><a:themeElements><a:clrScheme name="Office"/>'
        f'<a:fontScheme name="Office">{fonts("majorFont", major, major_ea, major_cs)}{fonts("minorFont", minor, minor_ea, minor_cs)}</a:fontScheme>'
        '<a:fmtScheme name="Office"/></a:themeElements></a:theme>'
    )


def hdr_xml(body, tag='hdr'):
    return f'{XML}<w:{tag} {W}>{body}</w:{tag}>'


def footnotes_xml(notes):
    """notes: list of body XML (paragraphs) for ids 1..n; Word's separator notes are added"""
    sep = (
        '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
    )
    body = ''.join(f'<w:footnote w:id="{i}">{xml}</w:footnote>' for i, xml in enumerate(notes, 1))
    return f'{XML}<w:footnotes {W} {R}>{sep}{body}</w:footnotes>'


def footnote_ref(i, rpr=''):
    return f'<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="{i}"/></w:r>'


PNG_1X1 = base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==')


def inline_pic(i, cx, cy, effect_b=0, name='pic'):
    """wp:inline picture of cx x cy EMU over media image i (write(..., media=[...]))"""
    WP = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
    A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    PIC = 'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
    return (
        f'<w:r><w:drawing><wp:inline {WP} distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="{effect_b}"/><wp:docPr id="{i + 1}" name="{name}{i}"/>'
        f'<a:graphic {A}><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic {PIC}><pic:nvPicPr><pic:cNvPr id="0" name="{name}{i}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="rIdImg{i}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
    )


def write(out, body, styles, settings=None, font_table=None, sect=LETTER, numbering=None, theme=None, sect_pr=None, hf=None, footnotes=None, media=None):
    """hf: {part name: xml (hdr_xml)}; a sectPr references it as r:id="rId<name without .xml>";
    footnotes: footnotes_xml(...) (body runs reference ids through footnote_ref);
    media: list of PNG bytes, referenced by inline_pic(i) as rIdImg<i>"""
    tail = sect_pr if sect_pr is not None else f'<w:sectPr>{sect}</w:sectPr>'
    doc = f'{XML}<w:document {W} {R}><w:body>{body}{tail}</w:body></w:document>'
    overrides = [
        ('/word/document.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'),
    ]
    rels = []
    parts = {'word/document.xml': doc}
    if styles is not None:
        overrides.append(('/word/styles.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml'))
        rels.append(('rId1', 'styles', 'styles.xml'))
        parts['word/styles.xml'] = styles
    if numbering is not None:
        overrides.append(('/word/numbering.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml'))
        rels.append(('rId4', 'numbering', 'numbering.xml'))
        parts['word/numbering.xml'] = numbering
    if theme is not None:
        overrides.append(('/word/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'))
        rels.append(('rId5', 'theme', 'theme/theme1.xml'))
        parts['word/theme/theme1.xml'] = theme
    if settings is not None:
        overrides.append(('/word/settings.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'))
        rels.append(('rId2', 'settings', 'settings.xml'))
        parts['word/settings.xml'] = settings
    for name, xml in (hf or {}).items():
        kind = 'header' if name.startswith('header') else 'footer'
        overrides.append((f'/word/{name}', f'application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml'))
        rels.append((f'rId{name[:-4]}', kind, name))
        parts[f'word/{name}'] = xml
    if footnotes is not None:
        overrides.append(('/word/footnotes.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml'))
        rels.append(('rId6', 'footnotes', 'footnotes.xml'))
        parts['word/footnotes.xml'] = footnotes
    if font_table is not None:
        overrides.append(('/word/fontTable.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml'))
        rels.append(('rId3', 'fontTable', 'fontTable.xml'))
        parts['word/fontTable.xml'] = font_table
    binary = {}
    for i, png in enumerate(media or []):
        rels.append((f'rIdImg{i}', 'image', f'media/image{i}.png'))
        binary[f'word/media/image{i}.png'] = png
    ct = (
        f'{XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        + ('<Default Extension="png" ContentType="image/png"/>' if binary else '')
        + ''.join(f'<Override PartName="{p}" ContentType="{c}"/>' for p, c in overrides)
        + '</Types>'
    )
    pkg_rels = (
        f'{XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>'
    )
    doc_rels = (
        f'{XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + ''.join(
            f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}" Target="{target}"/>'
            for rid, kind, target in rels
        )
        + '</Relationships>'
    )
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', ct)
        z.writestr('_rels/.rels', pkg_rels)
        z.writestr('word/_rels/document.xml.rels', doc_rels)
        for name, xml in parts.items():
            z.writestr(name, xml)
        for name, data in binary.items():
            z.writestr(name, data)


def out_path():
    return sys.argv[1] if len(sys.argv) > 1 else 'probe.docx'


def tbl(rows, width=6000, border_sz=4, cell_mar='', extra=''):
    """one-column fixed table; rows: list of cell content XML (paragraphs)"""
    bord = ''.join(
        f'<w:{s} w:val="single" w:sz="{border_sz}" w:space="0" w:color="000000"/>'
        for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')
    )
    trs = ''.join(
        f'<w:tr><w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>{c}</w:tc></w:tr>' for c in rows
    )
    return (
        f'<w:tbl><w:tblPr>{extra}<w:tblW w:w="{width}" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
        f'<w:tblBorders>{bord}</w:tblBorders>{cell_mar}</w:tblPr><w:tblGrid><w:gridCol w:w="{width}"/></w:tblGrid>{trs}</w:tbl>'
    )


def textbox(content, cx=2743200, cy=914400, ins=(91440, 45720, 91440, 45720), id=1, border=True):
    """inline wps text box (Word 2010+ reads the mc:Choice); a thin border so a PDF reader sees the box"""
    ns = (
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
    )
    ln = (
        '<a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>'
        if border
        else '<a:ln><a:noFill/></a:ln>'
    )
    l, t, r, b = ins
    return (
        f'<w:p><w:r><mc:AlternateContent {ns}><mc:Choice Requires="wps"><w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:docPr id="{id}" name="Text Box {id}"/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr txBox="1"/>'
        f'<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>{ln}</wps:spPr>'
        f'<wps:txbx><w:txbxContent>{content}</w:txbxContent></wps:txbx>'
        f'<wps:bodyPr wrap="square" lIns="{l}" tIns="{t}" rIns="{r}" bIns="{b}" anchor="t"><a:noAutofit/></wps:bodyPr>'
        '</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice></mc:AlternateContent></w:r></w:p>'
    )
