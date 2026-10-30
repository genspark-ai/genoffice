import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Arabic-only cell paragraphs in a cs face: Word paces the row by that face's own
# factor (Arial 1.15, Sakkal Majalla 1.40), never by the ascii Calibri strut.
AR = 'عنوان التقرير السنوي'
def cs(face, sz):
    return (f'<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="{face}"/>'
            f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/><w:rtl/></w:rPr>')
rows = [
    para(AR, ppr='<w:bidi/>', run_rpr=cs('Arial', 24)),
    para(AR, ppr='<w:bidi/>', run_rpr=cs('Sakkal Majalla', 28)),
    para('Latin control line', run_rpr=rpr('Calibri', 24)),
]
body = para('lead', run_rpr=rpr('Calibri', 24)) + tbl(rows) + para('tail', run_rpr=rpr('Calibri', 24))
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
