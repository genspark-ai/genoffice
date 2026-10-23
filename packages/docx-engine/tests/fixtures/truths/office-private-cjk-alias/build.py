import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
D = '等线'
M = '新細明體'
body = (para('文字文字 DengXian 11', run_rpr=rpr('Calibri', 22, east_asia=D))
        + para('文字文字 PMingLiU 12', run_rpr=rpr('Calibri', 24, east_asia=M)))
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22, east_asia=D)), settings_xml('<w:useFELayout/>'))
