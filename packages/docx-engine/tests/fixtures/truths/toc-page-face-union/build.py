"""TOC lines mixing a declared title face with an inherited page-number face (Word probe 2026-11-07).
Builds probe.docx; convert with Word and measure baseline pitches per marker group."""
import sys, zipfile
from pathlib import Path
from xml.sax.saxutils import escape
OUT = Path(__file__).resolve().parent
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CG = 'Century Gothic'
def rfonts(f): return f'<w:rFonts w:ascii="{f}" w:hAnsi="{f}"/>' if f else ''
def toc(title_font, tab_font, page_font, anchor):
    return ('<w:p><w:pPr><w:pStyle w:val="TM1"/><w:rPr><w:sz w:val="24"/></w:rPr></w:pPr>'
      f'<w:r><w:rPr>{rfonts(title_font)}<w:noProof/></w:rPr><w:t>Article 1 : Objet DU MARCHE</w:t></w:r>'
      f'<w:r><w:rPr>{rfonts(tab_font)}<w:noProof/></w:rPr><w:tab/></w:r>'
      f'<w:r><w:rPr>{rfonts(page_font)}<w:noProof/></w:rPr><w:t>5</w:t></w:r>'
      '</w:p>')
def marker(name): return f'<w:p><w:pPr><w:spacing w:after="0" w:line="276" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr><w:t>#{name}#</w:t></w:r></w:p>'
variants = [('A2_cg_inh_inh', CG, None, None), ('B2_cg_cg_cg', CG, CG, CG), ('B3_cg_inh_cg', CG, None, CG), ('D2_cg_cg_georgia', CG, CG, 'Georgia'), ('E2_cal_cal_cg', None, None, CG), ('C2_inh_inh_inh', None, None, None)]
body=[]; k=0
# a bookmark target so PAGEREF has something (Word keeps the cached "5" anyway)
pass
for name, tf, tabf, pf in variants:
    body.append(marker(name))
    for _ in range(6): body.append(toc(tf, tabf, pf, '_Toc1'))
body.append(marker('END'))
sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
document = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(body)}{sect}</w:body></w:document>'
styles = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="fr-FR" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="TM1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:tabs><w:tab w:val="left" w:pos="1320"/><w:tab w:val="right" w:leader="dot" w:pos="9214"/></w:tabs><w:spacing w:after="0"/><w:ind w:right="-284"/><w:contextualSpacing/></w:pPr><w:rPr><w:rFonts w:cs="Times New Roman"/><w:bCs/><w:caps/><w:sz w:val="20"/><w:szCs w:val="24"/></w:rPr></w:style><w:style w:type="character" w:styleId="Lienhypertexte"><w:name w:val="Hyperlink"/><w:uiPriority w:val="99"/><w:rPr><w:color w:val="000080"/><w:u w:val="single"/></w:rPr></w:style></w:styles>'''
settings = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'
ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
drels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'
docx = OUT / 'probe.docx'
with zipfile.ZipFile(docx, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', ct); z.writestr('_rels/.rels', rels)
    z.writestr('word/document.xml', document); z.writestr('word/styles.xml', styles)
    z.writestr('word/settings.xml', settings); z.writestr('word/_rels/document.xml.rels', drels)
print('built', docx)
