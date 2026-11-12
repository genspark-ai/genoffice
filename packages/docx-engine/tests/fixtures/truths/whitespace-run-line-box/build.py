"""Whitespace-only runs and the paragraph mark never size a line that holds ink (Word probe 2026-11-08).
Builds probe.docx; convert with Word and measure baseline pitches per marker group."""
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape
OUT = Path(__file__).resolve().parent
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
TNR = 'Times New Roman'
def rpr(font=None, sz=None):
    s = '<w:rPr>'
    if font: s += f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
    if sz: s += f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
    return s + '</w:rPr>'
def run(text, font=None, sz=None, kind='t', pre=''):
    if kind == 't': body = f'<w:t xml:space="preserve">{escape(text)}</w:t>'
    elif kind == 'tab': body = '<w:tab/>'
    else: body = ''
    return f'<w:r>{rpr(font, sz)}{pre}{body}</w:r>'
def para(runs, mark_rpr='', ppr_extra=''):
    return f'<w:p><w:pPr>{ppr_extra}{mark_rpr}</w:pPr>{"".join(runs)}</w:p>'
def marker(name): return para([run(f'#{name}#', 'Arial', 20)])
T1 = 'The night shift at the port was always noisy, but'
T2 = 'tonight the factory felt different.'
LONG = 'Machines roared, conveyor belts rattled, forklifts hummed, and distant horns echoed through the heavy coastal air while the workers sensed a quiet tension. '
variants = [
  ('A_tnr12_ctrl', lambda: [run(T1, TNR, 24), run(' ', TNR, 24), run(T2, TNR, 24)]),
  ('B_tnr12_space_cal11', lambda: [run(T1, TNR, 24), run(' ', None, 22), run(T2, TNR, 24)]),
  ('C_tnr12_space_cal14', lambda: [run(T1, TNR, 24), run(' ', None, 28), run(T2, TNR, 24)]),
  ('D_tnr12_mark_cal14', lambda: [run(T1 + ' ' + T2, TNR, 24)], rpr(None, 28)),
  ('E_tnr12_space_cal14_mark_cal14', lambda: [run(T1, TNR, 24), run(' ', None, 28), run(T2, TNR, 24)], rpr(None, 28)),
  ('F_tnr12_tab_cal14', lambda: [run(T1, TNR, 24), run('', None, 28, 'tab'), run(T2, TNR, 24)]),
  ('G_tnr12_nbsp_cal14', lambda: [run(T1, TNR, 24), run(' ', None, 28), run(T2, TNR, 24)]),
  ('K_tnr12_trailing_space_cal14', lambda: [run(T1 + ' ' + T2, TNR, 24), run(' ', None, 28)]),
  ('L_tnr12_mark_cal14_inherit_space', lambda: [run(T1, TNR, 24), run(' ', None, None), run(T2, TNR, 24)], rpr(None, 28)),
  ('J_tnr12_wrapped_space_cal14', lambda: [run(LONG, TNR, 24), run(' ', None, 28), run(LONG, TNR, 24), run(' ', None, 28), run(LONG, TNR, 24)]),
]
body = []
for v in variants:
    name, mk = v[0], v[1]
    mark = v[2] if len(v) > 2 else ''
    body.append(marker(name))
    for _ in range(6): body.append(para(mk(), mark))
# space-only paragraphs: mark sz 2, ' ' run inheriting 11pt
body.append(marker('H_spaceonly_mark2_run11'))
for _ in range(6): body.append(para([run(' ', None, None)], '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>'))
body.append(marker('I_spaceonly_mark2_run11_lrpb'))
for _ in range(6): body.append(para([run(' ', None, None, 't', '<w:lastRenderedPageBreak/>')], '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>'))
body.append(marker('M_spaceonly_mark2_run11_bookmark'))
for i in range(6): body.append(f'<w:p><w:pPr><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr><w:bookmarkStart w:id="{i}" w:name="_bm{i}"/><w:r><w:t xml:space="preserve"> </w:t></w:r><w:bookmarkEnd w:id="{i}"/></w:p>')
body.append(marker('N_empty_mark2'))
for _ in range(6): body.append(para([], '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>'))
body.append(marker('END'))
sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
document = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(body)}{sect}</w:body></w:document>'
styles = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style></w:styles>'''
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
