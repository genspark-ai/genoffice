"""probe.docx: Calibri 11 Normal (docDefaults after=0 line=240); page 1 two Calibri
control lines, then one page per face (Times New Roman, Georgia, Arial, Cambria):
declared-face runs first on the page, Calibri marker, declared-face runs mid page,
Calibri marker; then TNR runs with a TNR paragraph mark, then TNR runs at line=360."""
import zipfile
from xml.sax.saxutils import escape

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
      '</Types>')
RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>')
DOCRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
           '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
           '</Relationships>')
STYLES = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}>'
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>'
          '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
          '</w:styles>')


def rpr(f):
    return f'<w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}"/>' if f else ''


def para(text, font=None, mark=None, pb=False, line=240):
    pbx = '<w:pageBreakBefore/>' if pb else ''
    m = f'<w:rPr>{rpr(mark)}</w:rPr>' if mark else ''
    return (f'<w:p><w:pPr>{pbx}<w:spacing w:before="0" w:after="0" w:line="{line}" w:lineRule="auto"/>{m}</w:pPr>'
            f'<w:r><w:rPr>{rpr(font)}</w:rPr><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>')


body = [para('Hxg Calibri control line one'), para('Hxg Calibri control line two')]
for f in ['Times New Roman', 'Georgia', 'Arial', 'Cambria']:
    body += [para(f'Hxg {f} runs first on page', font=f, pb=True), para('Hxg Calibri marker after'),
             para(f'Hxg {f} runs mid page', font=f), para('Hxg Calibri marker end')]
body += [para('Hxg Times runs and mark first on page', font='Times New Roman', mark='Times New Roman', pb=True),
         para('Hxg Calibri marker after'),
         para('Hxg Times runs line 360 first on page', font='Times New Roman', pb=True, line=360),
         para('Hxg Calibri marker after')]
doc = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>' + ''.join(body) +
       '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
       'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>')
with zipfile.ZipFile('probe.docx', 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', CT)
    z.writestr('_rels/.rels', RELS)
    z.writestr('word/_rels/document.xml.rels', DOCRELS)
    z.writestr('word/document.xml', doc)
    z.writestr('word/styles.xml', STYLES)
