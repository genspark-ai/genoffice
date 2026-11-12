"""Batched Word probe for the six legacy line-factor rows (word-truths 2026-11-09): one page per
name, an Arial 12 marker paragraph then 9 br-joined 12 pt lines; fontTable entries carry the
PANOSE / charset the corpora declare. Usage: python3 build.py probe.docx"""
import sys, zipfile
from xml.sax.saxutils import escape
LATIN = "Hamburgefonstiv 123 quick brown fox"
ZH = "永和九年岁在癸丑暮春之初"
# marker, font, text, fontTable props
ROWS = [
    ("Calibri", "Calibri", LATIN, None),
    ("Times New Roman", "Times New Roman", LATIN, None),
    ("Arial", "Arial", LATIN, None),
    ("Times", "Times", LATIN, {"panose1": "02020603050405020304", "charset": "00", "family": "roman", "pitch": "variable"}),
    ("Helvetica", "Helvetica", LATIN, {"panose1": "020B0604020202020204", "charset": "00", "family": "swiss", "pitch": "variable"}),
    ("Helvetica 55 Roman", "Helvetica 55 Roman", LATIN, {"panose1": "020B0604020202020204", "charset": "00", "family": "swiss", "pitch": "variable"}),
    ("Arial Unicode MS", "Arial Unicode MS", LATIN, {"panose1": "020B0604020202020204", "charset": "00", "family": "swiss", "pitch": "variable"}),
    ("Arial Unicode MS|zh", "Arial Unicode MS", ZH, None),
    ("Georgia", "Georgia", LATIN, {"panose1": "02040502050405020303", "charset": "00", "family": "roman", "pitch": "variable"}),
    ("Segoe Print", "Segoe Print", LATIN, {"panose1": "02000800000000000000", "charset": "00", "family": "auto", "pitch": "variable"}),
    ("Segoe Script", "Segoe Script", LATIN, {"panose1": "020B0804020000000003", "charset": "00", "family": "swiss", "pitch": "variable"}),
    ("Segoe UI Emoji", "Segoe UI Emoji", LATIN, {"panose1": "020B0502040204020203", "charset": "00", "family": "swiss", "pitch": "variable"}),
    ("SimKai", "SimKai", LATIN, {"panose1": "02010609030101010101", "charset": "86", "family": "auto", "pitch": "variable"}),
    ("SimKai|zh", "SimKai", ZH, None),
    ("BiauKai", "BiauKai", LATIN, {"panose1": "02000500000000000000", "charset": "88", "family": "auto", "pitch": "variable"}),
    ("BiauKai|zh", "BiauKai", ZH, None),
    ("Kaiu", "Kaiu", LATIN, {"panose1": "02000500000000000000", "charset": "88", "family": "auto", "pitch": "variable"}),
    ("Kaiu|zh", "Kaiu", ZH, None),
]
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
def marker(m, pb):
    r = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="24"/>'
    return (f'<w:p><w:pPr>{"<w:pageBreakBefore/>" if pb else ""}<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr><w:t>{escape(f"#FONT:{m}#")}</w:t></w:r></w:p>')
def probe(font, text):
    r = f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}" w:cs="{font}"/><w:sz w:val="24"/><w:szCs w:val="24"/>'
    t = f'<w:t>{escape(text)}</w:t>'
    return (f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:rPr>{r}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{r}</w:rPr>{t}{("<w:br/>" + t) * 8}</w:r></w:p>')
body = "".join(marker(m, i > 0) + probe(f, t) for i, (m, f, t, _) in enumerate(ROWS))
doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>'
seen = {}
for _, f, _, props in ROWS:
    if props and f not in seen: seen[f] = props
fonts = "".join(f'<w:font w:name="{n}">' + "".join(f'<w:{k} w:val="{v}"/>' for k, v in p.items()) + '</w:font>' for n, p in seen.items())
font_table = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts {W}>{fonts}</w:fonts>'
styles = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>'
settings = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'
CT = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/></Types>'
RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
DRELS = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{R}styles" Target="styles.xml"/><Relationship Id="rId2" Type="{R}settings" Target="settings.xml"/><Relationship Id="rId3" Type="{R}fontTable" Target="fontTable.xml"/></Relationships>'
with zipfile.ZipFile(sys.argv[1], "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", CT); z.writestr("_rels/.rels", RELS); z.writestr("word/_rels/document.xml.rels", DRELS)
    z.writestr("word/document.xml", doc); z.writestr("word/styles.xml", styles); z.writestr("word/settings.xml", settings); z.writestr("word/fontTable.xml", font_table)
