import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
MG = '맑은 고딕'
H = '참조 문서 계약 범위 재확인'
SP = spacing(after=0, line=240, rule='auto')
cell_hm = '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:hideMark/></w:tcPr><w:p><w:pPr>' + SP + rpr('Calibri', 21, east_asia=MG) + '</w:pPr></w:p></w:tc>'
cell_pl = '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:pPr>' + SP + rpr('Calibri', 21, east_asia=MG) + '</w:pPr></w:p></w:tc>'
def tbl(cell):
    return f'<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="216"/></w:trPr>{cell}</w:tr></w:tbl>'
body = (para([H, H, H], ppr=SP, run_rpr=rpr('Calibri', 22, east_asia=MG))
        + para(None, ppr=SP, run_rpr=rpr('Calibri', 21, east_asia=MG), mark_rpr=rpr(MG, 21, east_asia=MG))
        + tbl(cell_hm) + para('x', ppr=SP) + tbl(cell_pl) + para('y', ppr=SP))
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22, east_asia=MG)), settings_xml('<w:useFELayout/>'), sect=A4)
