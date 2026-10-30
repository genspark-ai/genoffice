import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Empty cell paragraphs sized by their paragraph mark: a SimSun 12 mark is a SimSun
# line (1.30), a Calibri 12 mark a Calibri line (1.22), no mark rPr the body line.
def mark(face, sz):
    return (f'<w:rPr><w:rFonts w:ascii="{face}" w:hAnsi="{face}" w:eastAsia="{face}" w:cs="{face}"/>'
            f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>')
rows = [
    para(None, mark_rpr=mark('SimSun', 24)),
    para(None, mark_rpr=mark('Calibri', 24)),
    para(None),
    para('text control', run_rpr=rpr('Calibri', 22)),
]
body = para('lead') + tbl(rows) + para('tail')
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
