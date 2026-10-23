import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
ar = rpr('Arial', 24)
exact = spacing(before=0, after=0, line=240, rule='exact')
sect_para = f'<w:p><w:pPr>{exact}<w:sectPr>{LETTER}</w:sectPr>{ar}</w:pPr></w:p>'
body = ''.join([
    para('S-PREV', ppr=exact, run_rpr=ar),
    sect_para,
    para('S-TARGET after the section break', ppr=exact, run_rpr=ar),
    para('S-NEXT', ppr=exact, run_rpr=ar),
])
write(out_path(), body, styles_xml(doc_defaults('Arial', 24)), settings_xml(), sect='<w:type w:val="continuous"/>' + LETTER)
