import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
body = para('* * * * *', run_rpr=rpr('Arial MT Black', 24))
ft = font_table_xml([('Arial MT Black', {'altName': 'Courier New', 'charset': '00', 'family': 'swiss', 'pitch': 'variable'})])
write(out_path(), body, styles_xml(doc_defaults()), settings_xml(), font_table=ft)
