import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Paragraphs made of inline pictures only (one, two, three in a row, two of
# different heights), the same with trailing text, and under a 1.15 multiple;
# a Calibri 11 control line follows each so the Word PDF gives the paragraph height.
PT = 12700
pic = lambda h, w=120, b=0: inline_pic(0, w * PT, h * PT, b)
ctrl = lambda k: para(f'control {k}')
body = ctrl(0)
body += f'<w:p>{pic(60)}</w:p>' + ctrl(1)
body += f'<w:p>{pic(60)}{pic(60)}</w:p>' + ctrl(2)
body += f'<w:p>{pic(60)}{pic(60)}{pic(60)}</w:p>' + ctrl(3)
body += f'<w:p>{pic(60)}{pic(40)}</w:p>' + ctrl(4)
body += f'<w:p>{pic(60)}{pic(60)}<w:r><w:t xml:space="preserve"> tail text</w:t></w:r></w:p>' + ctrl(5)
body += f'<w:p>{pic(60)}<w:r><w:t xml:space="preserve"> tail text</w:t></w:r></w:p>' + ctrl(6)
body += f'<w:p><w:pPr>{spacing(line=276, rule="auto")}</w:pPr>{pic(60)}{pic(60)}</w:p>' + ctrl(7)
body += f'<w:p><w:pPr>{spacing(line=276, rule="auto")}</w:pPr>{pic(60)}</w:p>' + ctrl(8)
body += f'<w:p>{pic(60, b=8 * PT)}{pic(60, b=8 * PT)}</w:p>' + ctrl(9)
body += f'<w:p><w:pPr>{spacing(after=160, line=259, rule="auto")}</w:pPr>{pic(60)}{pic(60)}</w:p>' + ctrl(10)
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml(), media=[PNG_1X1])
