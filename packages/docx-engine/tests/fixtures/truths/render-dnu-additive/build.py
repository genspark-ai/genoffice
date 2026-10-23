import zipfile
W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
RPR='<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/>'
def p(text, ppr=''):
    r = f'<w:r><w:rPr>{RPR}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>' if text is not None else ''
    return f'<w:p><w:pPr>{ppr}<w:rPr>{RPR}</w:rPr></w:pPr>{r}</w:p>'
EXACT='<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>'
def sp(b,a,extra=''): return f'<w:spacing w:before="{b}" w:after="{a}" w:line="240" w:lineRule="exact"{extra}/>'
def filler(n,tag): return ''.join(p(f'{tag} filler {i+1}', EXACT) for i in range(n))
def brk(): return '<w:p><w:pPr>'+EXACT+'</w:pPr><w:r><w:br w:type="page"/></w:r></w:p>'
cases=[]
cases.append(('M', p('M-PREV after200', sp(0,200)) + p('M-TARGET before220', sp(220,0)) + p('M-NEXT', sp(0,0))))
cases.append(('N', p('N-PREV after200', sp(0,200)) + p('N-TARGET before0', sp(0,0)) + p('N-NEXT', sp(0,0))))
cases.append(('H', p('H-PREV after200', sp(0,200)) + p('H-TARGET beforeAuto', '<w:spacing w:beforeAutospacing="1" w:after="0" w:line="240" w:lineRule="exact"/>') + p('H-NEXT', sp(0,0))))
cases.append(('T', filler(54,'T') + p('T-TARGET before360', sp(360,0)) + p('T-NEXT', sp(0,0))))
cases.append(('G', filler(53,'G') + p('G-PREV after240', sp(0,240)) + p('G-TARGET before360', sp(360,0)) + p('G-NEXT', sp(0,0))))
body=''
for i,(tag,xml) in enumerate(cases):
    if i: body+=brk()
    body+=xml
body+='<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
doc=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}</w:body></w:document>'
def build(flag, out):
    f='<w:doNotUseHTMLParagraphAutoSpacing/>' if flag else ''
    settings=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat>{f}<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'
    styles=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>{RPR}</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'
    ct='<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
    rels='<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    drels='<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml',ct); z.writestr('_rels/.rels',rels)
        z.writestr('word/_rels/document.xml.rels',drels); z.writestr('word/document.xml',doc)
        z.writestr('word/styles.xml',styles); z.writestr('word/settings.xml',settings)
import sys
build(True, sys.argv[1] if len(sys.argv) > 1 else 'probe.docx')
