import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
faces = ['Roboto', 'Lato', 'Montserrat', 'Open Sans', 'Inter', 'Trebuchet MS', 'Century Gothic', 'Helvetica Neue', 'Nirmala UI', 'Verdana', 'Georgia Pro']
body = ''.join(para([f'{f} line one', 'line two'], run_rpr=rpr(f, 24)) for f in faces)
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
