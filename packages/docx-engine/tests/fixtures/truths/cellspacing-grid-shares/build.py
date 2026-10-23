import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
body = para('A')
write(out_path(), body, styles_xml(doc_defaults('Times New Roman', 24)), settings_xml())
