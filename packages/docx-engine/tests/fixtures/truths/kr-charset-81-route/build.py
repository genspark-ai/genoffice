import zipfile, sys
out = sys.argv[1]
# (name, panose, charset, family)  None = omit attribute; 'SKIP' = no fontTable entry
FONTS = [
 ("함초롬바탕", "02030604000101010101", "81", "modern"),   # HCR Batang as in sample
 ("다람쥐서체1", "02030604000101010101", "81", "roman"),
 ("다람쥐서체2", "02030604000101010101", "00", "roman"),
 ("다람쥐서체3", None, "81", "auto"),
 ("다람쥐서체4", "02030604000101010101", None, "roman"),
 ("다람쥐서체5", "SKIP", None, None),
 ("Squirrel Serif Probe", "02030604000101010101", "81", "roman"),
 ("Squirrel Serif Probe2", "02030604000101010101", "00", "roman"),
 ("함초롬돋움", "020B0604000101010101", "81", "modern"),  # HCR Dotum
 ("다람쥐서체9", "02030604000101010101", "81", "modern"),
 ("Squirrel Serif Probe3", None, "81", "auto"),
]
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
def para(i, name):
    rpr = f'<w:rPr><w:rFonts w:ascii="{name}" w:hAnsi="{name}" w:eastAsia="{name}" w:cs="{name}"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'
    lines = [f"P{i} 한글 줄 하나", "한글 줄 둘", "한글 줄 셋", f"P{i} Latin line one", "Latin line two"]
    runs = ''.join(f'<w:r>{rpr}<w:t xml:space="preserve">{t}</w:t></w:r>' + (f'<w:r>{rpr}<w:br/></w:r>' if k < len(lines)-1 else '') for k, t in enumerate(lines))
    return f'<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>{runs}</w:p>'
sep = '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr><w:t>---- Calibri separator ----</w:t></w:r></w:p>'
body = ''.join(para(i, f[0]) + sep for i, f in enumerate(FONTS))
doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>'
def fent(name, panose, charset, family):
    if panose == 'SKIP': return ''
    s = f'<w:font w:name="{name}">'
    if panose: s += f'<w:panose1 w:val="{panose}"/>'
    if charset: s += f'<w:charset w:val="{charset}"/>'
    if family: s += f'<w:family w:val="{family}"/>'
    s += '<w:pitch w:val="variable"/></w:font>'
    return s
ft = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts {W}>{"".join(fent(*f) for f in FONTS)}<w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/></w:font></w:fonts>'
styles = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ko-KR" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'
settings = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat><w:balanceSingleByteDoubleByteWidth/><w:useFELayout/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'
ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
drels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', ct); z.writestr('_rels/.rels', rels); z.writestr('word/_rels/document.xml.rels', drels)
    z.writestr('word/document.xml', doc); z.writestr('word/styles.xml', styles); z.writestr('word/fontTable.xml', ft); z.writestr('word/settings.xml', settings)
print(out)
