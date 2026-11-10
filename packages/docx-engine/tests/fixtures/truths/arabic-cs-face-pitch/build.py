import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
AR = 'صف خلق خود كمثل الشمس'
def p(sz, line, cs):
    return f'<w:p><w:pPr><w:bidi/>{spacing(after=0, line=line, rule="auto")}</w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="{cs}"/><w:rtl/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr><w:t>{AR}</w:t></w:r></w:p>'
body = p(24, 240, 'Arial') + p(28, 240, 'Arial') + p(24, 276, 'Arial') + p(28, 276, 'Arial')
write(out_path(), body, styles_xml(doc_defaults('Calibri', 24)), settings_xml())
