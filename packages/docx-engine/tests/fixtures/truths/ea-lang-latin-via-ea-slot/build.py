import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Google Docs export of a Korean document: Arial in every slot, w:lang w:val="ko"
dd = ('<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:cs="Arial" w:eastAsia="Arial" w:hAnsi="Arial"/>'
      '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="ko"/></w:rPr></w:rPrDefault>'
      f'<w:pPrDefault><w:pPr>{spacing(after=0, line=240, rule="auto")}</w:pPr></w:pPrDefault></w:docDefaults>')
LAT = 'Shaker Tarima 0123 +----+ Journal'
KO = '한국형 의과대학 K-Team Science'
def every(f): return rpr(f, east_asia=f)
body = (para('A1 ' + LAT)
        + para('A2 ' + KO)
        + para('A3 ' + LAT, run_rpr=every('Courier New'))
        + para('A4 ' + KO, run_rpr=every('Courier New'))
        + para('A5 ' + LAT, run_rpr=every('Arial Unicode MS'))
        + para('A6 ' + KO, run_rpr=every('Arial Unicode MS'))
        + para('A7 ' + LAT, run_rpr=rpr('Courier New', east_asia='Courier New', extra='<w:lang w:val="en-US"/>'))
        + para('A8 ' + LAT, run_rpr=rpr(extra='<w:lang w:val="en-US"/>'))
        + para('A9 ' + LAT, run_rpr=rpr('Courier New', east_asia='Batang'))
        + para('A10 ' + LAT, run_rpr=every('Gungsuh'))
        + para('A11 ' + KO, run_rpr=every('Gungsuh')))
ft = font_table_xml([('Arial', {}), ('Arial Unicode MS', {}), ('Courier New', {}), ('Gungsuh', {})])
write(out_path(), body, styles_xml(dd), settings_xml(), font_table=ft, sect=A4)
