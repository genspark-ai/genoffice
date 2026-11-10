import zipfile, sys, os
W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
def p(text, ppr='', rpr='<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/>'):
    r = f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>' if text is not None else ''
    return f'<w:p><w:pPr>{ppr}<w:rPr>{rpr}</w:rPr></w:pPr>{r}</w:p>'
EXACT='<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>'
AUTO='<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
def filler(n, tag):
    return ''.join(p(f'{tag} filler {i+1}', EXACT) for i in range(n))
def brk():
    return '<w:p><w:pPr>'+EXACT+'</w:pPr><w:r><w:br w:type="page"/></w:r></w:p>'
cases=[]
# E control: full page then plain paragraph
cases.append(('E', filler(54,'E') + p('E-TARGET plain', AUTO) + p('E-NEXT', AUTO)))
# A: full page then heading before=360
cases.append(('A', filler(54,'A') + p('A-TARGET before360', '<w:spacing w:before="360" w:after="0" w:line="240" w:lineRule="auto"/>') + p('A-NEXT', AUTO)))
# A2: 24pt left: line fits, before+line does not
cases.append(('A2', filler(52,'A2') + p('A2-TARGET before360', '<w:spacing w:before="360" w:after="0" w:line="240" w:lineRule="auto"/>') + p('A2-NEXT', AUTO)))
# A3: 36pt left: before+line fits (control for A2)
cases.append(('A3', filler(51,'A3') + p('A3-TARGET before360', '<w:spacing w:before="360" w:after="0" w:line="240" w:lineRule="auto"/>') + p('A3-NEXT', AUTO)))
# B: 53 fillers + exact-line para with after=200 (after overflows) then text
cases.append(('B', filler(53,'B') + p('B-LAST after200', '<w:spacing w:before="0" w:after="200" w:line="240" w:lineRule="exact"/>') + p('B-TARGET', AUTO) + p('B-NEXT', AUTO)))
# C: full page, then empty after=240, then MARK: empty leads p2
cases.append(('C', filler(54,'C') + p(None, '<w:spacing w:before="0" w:after="240" w:line="240" w:lineRule="exact"/>') + p('C-TARGET after empty', AUTO) + p('C-NEXT', AUTO)))
# D: keepNext chain pushed: 50 fillers (48pt left) + heading before360 keepNext + 6 exact lines paragraph
cases.append(('D', filler(50,'D') + p('D-TARGET before360 keepNext', '<w:keepNext/><w:spacing w:before="360" w:after="0" w:line="240" w:lineRule="auto"/>') + ''.join(p(f'D-body {i}', EXACT) for i in range(6)) + p('D-NEXT', AUTO)))
# F: widow/orphan pushed 3-line auto paragraph after a para with after=240 (like sample 012 p7)
cases.append(('F', filler(52,'F') + p('F-PREV after240', '<w:spacing w:before="0" w:after="240" w:line="240" w:lineRule="exact"/>') + p('F-TARGET line one '+'word '*120, AUTO) + p('F-NEXT', AUTO)))
# G: before=360 keepNext pushed, compat-independent? also the previous paragraph has after=240 -> check both drop
cases.append(('G', filler(53,'G') + p('G-PREV after240', '<w:spacing w:before="0" w:after="240" w:line="240" w:lineRule="exact"/>') + p('G-TARGET before360', '<w:spacing w:before="360" w:after="0" w:line="240" w:lineRule="auto"/>') + p('G-NEXT', AUTO)))
body=''
for i,(tag,xml) in enumerate(cases):
    if i: body+=brk()
    body+=xml
body+='<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
doc=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}</w:body></w:document>'
def build(compat, out):
    settings=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/></w:compat></w:settings>'
    styles=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'
    ct='<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
    rels='<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    drels='<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml',ct); z.writestr('_rels/.rels',rels)
        z.writestr('word/_rels/document.xml.rels',drels); z.writestr('word/document.xml',doc)
        z.writestr('word/styles.xml',styles); z.writestr('word/settings.xml',settings)
import sys
build(15, sys.argv[1] if len(sys.argv) > 1 else 'probe.docx')
