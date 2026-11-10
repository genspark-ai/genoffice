import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
AUTO = '<w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1" w:line="240" w:lineRule="auto"/>'
body = ''.join([
    para('plain 11pt', ppr=spacing(before=0, after=0, line=240, rule='auto')),
    para('auto spacing 11pt', ppr=AUTO, run_rpr=rpr('Calibri', 22)),
    para('auto spacing 17pt', ppr=AUTO, run_rpr=rpr('Calibri', 34)),
    para('plain 11pt', ppr=spacing(before=0, after=0, line=240, rule='auto')),
])
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
