import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
body = ''.join([
    para(' Symbol bullet run', run_rpr=rpr('Symbol', 24)),
    para(' Wingdings run', run_rpr=rpr('Wingdings', 24)),
    para(' Wingdings 2 run', run_rpr=rpr('Wingdings 2', 24)),
    para(' Wingdings 3 run', run_rpr=rpr('Wingdings 3', 24)),
    para(' Webdings run', run_rpr=rpr('Webdings', 24)),
])
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
