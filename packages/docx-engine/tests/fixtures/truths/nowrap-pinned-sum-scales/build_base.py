import zipfile, sys, os
OUT=sys.argv[1]
W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
def cell(w, text, nowrap=False, span=None):
    pr=f'<w:tcW w:w="{w}" w:type="dxa"/>'+(f'<w:gridSpan w:val="{span}"/>' if span else '')+('<w:noWrap/>' if nowrap else '')
    return f'<w:tc><w:tcPr>{pr}</w:tcPr><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p></w:tc>'
def table(label, grid, rows, layout=None, tblw='<w:tblW w:w="0" w:type="auto"/>'):
    tp='<w:tblPr>'+tblw+(f'<w:tblLayout w:type="{layout}"/>' if layout else '')+'<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders><w:tblCellMar><w:left w:w="15" w:type="dxa"/><w:right w:w="15" w:type="dxa"/></w:tblCellMar></w:tblPr>'
    g='<w:tblGrid>'+''.join(f'<w:gridCol w:w="{c}"/>' for c in grid)+'</w:tblGrid>'
    trs=''.join('<w:tr>'+''.join(r)+'</w:tr>' for r in rows)
    para=f'<w:p><w:pPr><w:spacing w:before="240" w:after="120" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t xml:space="preserve">{label}</w:t></w:r></w:p>'
    return para+'<w:tbl>'+tp+g+trs+'</w:tbl><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:p>'
long='Bendahara penerima menerima tagihan dari pajak PBB sebesar Rp 11.000.000 untuk tahun berjalan dan mencatatnya'
body=''
# P1: 056 mirror: autofit, grid 926+49+8051 = 9026, tcW 1530 span2 noWrap + 12915; text wider than 975 twips
body+=table('P1 autofit overdeclared noWrap wide', [926,49,8051], [[cell(1530,'02 Jan 2025',True,2),cell(12915,long)],[cell(1530,'03 Jan 2025',True,2),cell(12915,long)]])
# P2: same but narrow text in the noWrap cell
body+=table('P2 autofit overdeclared noWrap narrow', [926,49,8051], [[cell(1530,'02',True,2),cell(12915,long)],[cell(1530,'03',True,2),cell(12915,long)]])
# P3: fixed layout variant of P1
body+=table('P3 fixed overdeclared noWrap wide', [926,49,8051], [[cell(1530,'02 Jan 2025',True,2),cell(12915,long)],[cell(1530,'03 Jan 2025',True,2),cell(12915,long)]], layout='fixed')
# P4: autofit, tcW within grid (tcW == grid), noWrap wide text
body+=table('P4 autofit tcW=grid noWrap wide', [975,8051], [[cell(975,'02 Jan 2025',True),cell(8051,long)],[cell(975,'03 Jan 2025',True),cell(8051,long)]])
# P5: autofit, tcW sum == column exactly, noWrap wide text
body+=table('P5 autofit tcW sum=column noWrap wide', [975,8051], [[cell(975,'02 Jan 2025',True),cell(8051,long)],[cell(975,'03 Jan 2025',True),cell(8051,long)]], tblw='<w:tblW w:w="9026" w:type="dxa"/>')
# P6: P1 without noWrap (control: does the autofit grid stay 975?)
body+=table('P6 autofit overdeclared no noWrap', [926,49,8051], [[cell(1530,'02 Jan 2025',False,2),cell(12915,long)],[cell(1530,'03 Jan 2025',False,2),cell(12915,long)]])
# P7: autofit, over-declared tcW, noWrap, but grid itself wide: grid 1530+7496
body+=table('P7 autofit overdeclared noWrap grid1530', [1530,7496], [[cell(1530,'02 Jan 2025',True),cell(12915,long)],[cell(1530,'03 Jan 2025',True),cell(12915,long)]])
# P8: autofit over-declared, noWrap text wider than tcW too (1530 = 76pt; text ~ 110pt)
body+=table('P8 autofit overdeclared noWrap very wide', [926,49,8051], [[cell(1530,'02 January 2025 Thursday',True,2),cell(12915,long)],[cell(1530,'03 January 2025 Friday',True,2),cell(12915,long)]])
doc=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>'
styles=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="DengXian" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>'
settings=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'
ct='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
rels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
drels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'
with zipfile.ZipFile(OUT,'w',zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml',ct); z.writestr('_rels/.rels',rels); z.writestr('word/_rels/document.xml.rels',drels)
    z.writestr('word/document.xml',doc); z.writestr('word/styles.xml',styles); z.writestr('word/settings.xml',settings)
print(OUT)
