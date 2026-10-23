"""Minimal synthetic docx writer shared by the truth fixtures' build.py scripts.

Usage: python3 build.py <out.docx>  (from the fixture directory)
"""
import sys
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
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


def write(out, body, styles, settings=None, font_table=None, sect=LETTER, numbering=None, theme=None, sect_pr=None):
    tail = sect_pr if sect_pr is not None else f'<w:sectPr>{sect}</w:sectPr>'
    doc = f'{XML}<w:document {W}><w:body>{body}{tail}</w:body></w:document>'
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
    if font_table is not None:
        overrides.append(('/word/fontTable.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml'))
        rels.append(('rId3', 'fontTable', 'fontTable.xml'))
        parts['word/fontTable.xml'] = font_table
    ct = (
        f'{XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
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


def out_path():
    return sys.argv[1] if len(sys.argv) > 1 else 'probe.docx'
