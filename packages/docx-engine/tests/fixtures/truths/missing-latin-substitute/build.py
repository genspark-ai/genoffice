import zipfile, sys
out = sys.argv[1]
# (name, panose, charset, family); 'SKIP' panose = no fontTable entry
FONTS = {
 "Mont": ("00000700000000000000", "CC", "auto"),
 "Mont Probe Zero": ("00000700000000000000", "00", "auto"),
 "Squirrel Probe Cyr": (None, "CC", "auto"),
 "Squirrel Probe Lat": (None, "00", "auto"),
 "Zebra Probe Nocharset": (None, None, None),
 "Zebra Probe Noentry": ("SKIP", None, None),
 "Helvetica Neue": ("02000503000000020004", "00", "auto"),
 "Calibri": ("020F0502020204030204", "00", "swiss"),
}
CYR = ["Пример текста строка один", "Пример текста строка два", "Пример текста строка три"]
LAT = ["Latin sample line one", "Latin sample line two", "Latin sample line three"]
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
def para(label, name, lines, sz=24, bold=False):
    b = '<w:b/>' if bold else ''
    rpr = f'<w:rPr><w:rFonts w:ascii="{name}" w:hAnsi="{name}" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>{b}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
    ls = [f"{label} " + lines[0]] + lines[1:]
    runs = ''.join(f'<w:r>{rpr}<w:t xml:space="preserve">{t}</w:t></w:r>' + (f'<w:r>{rpr}<w:br/></w:r>' if k < len(ls)-1 else '') for k, t in enumerate(ls))
    return f'<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>{runs}</w:p>'
sep = para("SEP", "Calibri", ["---- separator ----"])
parts = []
i = 0
for name in ["Mont", "Mont Probe Zero", "Squirrel Probe Cyr", "Squirrel Probe Lat", "Zebra Probe Nocharset", "Zebra Probe Noentry"]:
    parts.append(para(f"P{i}C", name, CYR)); i += 1
    parts.append(para(f"P{i}L", name, LAT)); i += 1
    parts.append(sep)
parts.append(para(f"P{i}H12", "Helvetica Neue", LAT, 24)); i += 1
parts.append(para(f"P{i}H18B", "Helvetica Neue", LAT, 36, True)); i += 1
parts.append(para(f"P{i}H10", "Helvetica Neue", LAT, 20)); i += 1
parts.append(sep)
parts.append(para(f"P{i}CAL", "Calibri", CYR, 24)); i += 1
body = ''.join(parts)
doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>'
def fent(name, panose, charset, family):
    if panose == 'SKIP': return ''
    s = f'<w:font w:name="{name}">'
    if panose: s += f'<w:panose1 w:val="{panose}"/>'
    if charset: s += f'<w:charset w:val="{charset}"/>'
    if family: s += f'<w:family w:val="{family}"/>'
    s += '<w:pitch w:val="variable"/></w:font>'
    return s
ft = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts {W}>{"".join(fent(n, *v) for n, v in FONTS.items())}</w:fonts>'
styles = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'
settings = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'
ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
drels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', ct); z.writestr('_rels/.rels', rels); z.writestr('word/_rels/document.xml.rels', drels)
    z.writestr('word/document.xml', doc); z.writestr('word/styles.xml', styles); z.writestr('word/fontTable.xml', ft); z.writestr('word/settings.xml', settings)
print(out)
