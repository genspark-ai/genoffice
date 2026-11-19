import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
A = rpr('Arial', 24)
tg = ('<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:pPr>' + spacing(after=0, line=240, rule='auto') + '</w:pPr>'
      '<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>')
def cell(t, ppr=''):
    return f'<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>{para([t, t, t], ppr=ppr, run_rpr=A)}</w:tc>'
tbl = ('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>'
       '<w:tr>' + cell('no style') + cell('direct 480', spacing(line=480, rule='auto')) + '</w:tr></w:tbl>')
body = para('control', run_rpr=A) + tbl + para('after', run_rpr=A)
def build(out, link):
    settings = f'{XML}<w:settings {W}>' + ('<w:linkStyles/>' if link else '') + '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'
    write(out, body, styles_xml(doc_defaults('Arial', 24, ppr=spacing(after=0, line=240, rule='auto')), tg), settings, sect=A4)
build(out_path(), False)
build('probe-link.docx', True)
