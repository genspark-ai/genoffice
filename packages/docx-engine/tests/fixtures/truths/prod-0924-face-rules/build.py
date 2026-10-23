import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
NM = '나눔명조'
body = (para('Helvetica body 10.5', run_rpr=rpr('Helvetica', 21))
        + para('한글 본문', run_rpr=rpr(NM, 20, east_asia=NM)))
ft = font_table_xml([(NM, {'altName': '바탕', 'charset': '81', 'family': 'roman'}), ('Helvetica', {'charset': '00', 'family': 'swiss'})])
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml(), font_table=ft)
