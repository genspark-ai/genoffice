import zipfile, sys, os
import os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import out_path
OUT = out_path()
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
def p(text=None, after=0, sz=22, font='Calibri'):
    rpr = f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
    ppr = f'<w:pPr><w:spacing w:before="0" w:after="{after}" w:line="240" w:lineRule="auto"/>{rpr}</w:pPr>'
    run = f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>' if text is not None else ''
    return f'<w:p>{ppr}{run}</w:p>'
def tc(paras, hide, w=4000):
    hm = '<w:hideMark/>' if hide else ''
    return f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{hm}</w:tcPr>{"".join(paras)}</w:tc>'
def tbl(cells):
    b = ''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>' for s in ('top','left','bottom','right','insideH','insideV'))
    grid = ''.join('<w:gridCol w:w="4000"/>' for _ in cells)
    return (f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>{b}</w:tblBorders>'
            f'<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
            f'<w:tblGrid>{grid}</w:tblGrid><w:tr>{"".join(cells)}</w:tr></w:tbl>')
T = lambda **k: p('Abc', **k)
E = lambda **k: p(None, **k)
cases = [
 ('H_T',   [tc([T()], True)]),
 ('H_TE',  [tc([T(), E()], True)]),
 ('H_TEE', [tc([T(), E(), E()], True)]),
 ('H_ET',  [tc([E(), T()], True)]),
 ('N_TE',  [tc([T(), E()], False)]),
 ('N_T',   [tc([T()], False)]),
 ('H_TEa', [tc([T(), E(after=160)], True)]),
 ('H_TaE', [tc([T(after=160), E()], True)]),
 ('H_TaEa',[tc([T(after=160), E(after=160)], True)]),
 ('N_TaEa',[tc([T(after=160), E(after=160)], False)]),
 ('H_TET', [tc([T(), E(), T()], True)]),
 ('H_E',   [tc([E()], True)]),
 ('H_TE28',[tc([T(), E(sz=28)], True)]),
 ('H_TTE_TET', [tc([T(after=160), T(after=160), E(after=160)], True), tc([T(after=160), E(after=160), T(after=160)], True)]),
 ('H_TTE_T', [tc([T(after=160), T(after=160), E(after=160)], True), tc([T(after=160)], True)]),
]
body = ''
for name, cells in cases:
    body += p(f'MARK {name}', font='Arial') + tbl(cells)
body += p('END', font='Arial')
body += '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>'
doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}</w:body></w:document>'
styles = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:eastAsia="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>'
ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
drels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', ct); z.writestr('_rels/.rels', rels)
    z.writestr('word/document.xml', doc); z.writestr('word/styles.xml', styles); z.writestr('word/_rels/document.xml.rels', drels)
print(OUT)
