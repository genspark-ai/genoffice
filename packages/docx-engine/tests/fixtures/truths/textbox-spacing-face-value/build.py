import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Text-box paragraph spacing at face value: before 6pt / after 4pt between two
# Calibri 11 lines is a 19.43 pitch (13.43 + 6), auto spacing is Word's 14pt.
def p(text, **sp):
    return para(text, ppr=spacing(**sp), run_rpr=rpr('Calibri', 22))
content = (
    p('first paragraph', after=80)
    + p('second paragraph', before=120, after=80)
    + p('third paragraph', before=120)
    + para('auto spaced', ppr='<w:spacing w:beforeAutospacing="1" w:afterAutospacing="1"/>', run_rpr=rpr('Calibri', 22))
    + p('last paragraph')
)
body = para('lead') + textbox(content, cx=2743200, cy=1280160) + para('tail')
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
