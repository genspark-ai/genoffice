import zipfile, sys
W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
def bd(sz):
    if sz is None:
        return ''.join(f'<w:{s} w:val="nil"/>' for s in ['top','left','bottom','right','insideH','insideV'])
    return ''.join(f'<w:{s} w:val="single" w:sz="{sz}" w:space="0" w:color="auto"/>' for s in ['top','left','bottom','right','insideH','insideV'])
def tbl(label, trh, sz, mar, n=4, marb=None):
    marb = mar if marb is None else marb
    marx = f'<w:tblCellMar><w:top w:w="{mar}" w:type="dxa"/><w:left w:w="28" w:type="dxa"/><w:bottom w:w="{marb}" w:type="dxa"/><w:right w:w="28" w:type="dxa"/></w:tblCellMar>'
    rows=''.join(f'<w:tr><w:trPr><w:trHeight w:hRule="exact" w:val="{trh}"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>{label} r{i}</w:t></w:r></w:p></w:tc></w:tr>' for i in range(n))
    return f'<w:p><w:r><w:t>{label}</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblBorders>{bd(sz)}</w:tblBorders><w:tblLayout w:type="fixed"/>{marx}</w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>{rows}</w:tbl>'
body = ''.join([
 tbl('H sz8 top100 bot0 440', 440, 8, 100, 4, 0),
 tbl('I sz8 top0 bot100 440', 440, 8, 0, 4, 100),
 tbl('J sz8 top200 bot100 440', 440, 8, 200, 4, 100),
])
doc=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>'
styles=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'
ct='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
rels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
drels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
with zipfile.ZipFile(sys.argv[1],'w',zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml',ct); z.writestr('_rels/.rels',rels); z.writestr('word/document.xml',doc); z.writestr('word/styles.xml',styles); z.writestr('word/_rels/document.xml.rels',drels)
