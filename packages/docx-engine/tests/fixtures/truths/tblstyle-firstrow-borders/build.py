import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
C = rpr('Calibri', 22)
ts = ('<w:style w:type="table" w:styleId="GridTable1Light"><w:name w:val="Grid Table 1 Light"/><w:tblPr><w:tblBorders>'
      '<w:top w:val="single" w:sz="4" w:space="0" w:color="999999"/><w:left w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
      '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="999999"/><w:right w:val="single" w:sz="4" w:space="0" w:color="999999"/>'
      '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="999999"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="999999"/></w:tblBorders></w:tblPr>'
      '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr><w:tcPr><w:tcBorders><w:bottom w:val="single" w:sz="12" w:space="0" w:color="666666"/></w:tcBorders></w:tcPr></w:tblStylePr></w:style>')
def row(texts):
    return '<w:tr>' + ''.join(f'<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>{para(t, run_rpr=C)}</w:tc>' for t in texts) + '</w:tr>'
tbl = ('<w:tbl><w:tblPr><w:tblStyle w:val="GridTable1Light"/><w:tblW w:w="6000" w:type="dxa"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>'
       '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' + row(['Head A', 'Head B']) + row(['r1a', 'r1b']) + row(['r2a', 'r2b']) + '</w:tbl>')
body = para('before') + tbl + para('after')
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22), ts), settings_xml())
