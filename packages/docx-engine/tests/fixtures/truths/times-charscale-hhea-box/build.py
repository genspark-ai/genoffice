import zipfile,sys
OUT=sys.argv[1]
W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
def para(label, font, sz, line, w=None, n=6):
    rpr=f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}" w:cs="{font}"/>'+(f'<w:w w:val="{w}"/>' if w else '')+f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
    lines='<w:br/>'.join(f'<w:t xml:space="preserve">Hamburgefonstiv line {i} of {label} the quick brown fox</w:t>' for i in range(n))
    marker=f'<w:p><w:pPr><w:spacing w:before="240" w:after="120" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr><w:t>{label}</w:t></w:r></w:p>'
    return marker+f'<w:p><w:pPr><w:spacing w:line="{line}" w:lineRule="auto" w:before="0" w:after="0"/><w:jc w:val="both"/></w:pPr><w:r><w:rPr>{rpr}</w:rPr>{lines}</w:r></w:p>'
body=''
body+=para('TNRw90','Times New Roman',21,420,90)
body+=para('TNRw110','Times New Roman',21,420,110)
body+=para('TNRw200','Times New Roman',21,420,200)
body+=para('TNRw100','Times New Roman',21,420,100)
body+=para('TNRw101','Times New Roman',21,420,101)
body+=para('Arialw105','Arial',24,240,105)
body+=para('Arialw100','Arial',24,240)
body+=para('Calibriw105','Calibri',24,240,105)
body+=para('Calibriw100','Calibri',24,240)
body+=para('Cambriaw105','Cambria',24,240,105)
body+=para('Cambriaw100','Cambria',24,240)
body+=para('TNRw105s240','Times New Roman',24,240,105)
doc=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>'
styles=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:widowControl w:val="0"/><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>'
fonts=''.join(f'<w:font w:name="{n}"><w:altName w:val="{n}"/><w:charset w:val="0"/><w:family w:val="{f}"/><w:pitch w:val="variable"/></w:font>' for n,f in [('Liberation Serif','roman'),('Times New Roman','roman'),('Liberation Sans','swiss'),('Arial','swiss'),('Noto Sans JP','swiss')])
fonttable=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts {W}>{fonts}</w:fonts>'
settings=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'
ct='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/></Types>'
rels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
drels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/></Relationships>'
with zipfile.ZipFile(OUT,'w',zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml',ct); z.writestr('_rels/.rels',rels); z.writestr('word/_rels/document.xml.rels',drels)
    z.writestr('word/document.xml',doc); z.writestr('word/styles.xml',styles); z.writestr('word/settings.xml',settings); z.writestr('word/fontTable.xml',fonttable)
