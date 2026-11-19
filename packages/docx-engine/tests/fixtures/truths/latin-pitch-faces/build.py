import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
body = ''.join([
    para(['Calibri 11 line one', 'line two', 'line three'], run_rpr=rpr('Calibri', 22)),
    para(['Calibri 12 line one', 'line two', 'line three'], run_rpr=rpr('Calibri', 24)),
    para(['Arial 12 line one', 'line two', 'line three'], run_rpr=rpr('Arial', 24)),
    para(['Arial 10 line one', 'line two', 'line three'], run_rpr=rpr('Arial', 20)),
    para(['Times New Roman 12 line one', 'line two', 'line three'], run_rpr=rpr('Times New Roman', 24)),
    para(['Georgia 12 line one', 'line two', 'line three'], run_rpr=rpr('Georgia', 24)),
    para(['Cambria 12 line one', 'line two', 'line three'], run_rpr=rpr('Cambria', 24)),
    para(['Consolas 12 line one', 'line two', 'line three'], run_rpr=rpr('Consolas', 24)),
    para(['Courier New 12 line one', 'line two', 'line three'], run_rpr=rpr('Courier New', 24)),
    para(['Segoe UI 12 line one', 'line two', 'line three'], run_rpr=rpr('Segoe UI', 24)),
    para(['Aptos 12 line one', 'line two', 'line three'], run_rpr=rpr('Aptos', 24)),
])
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
