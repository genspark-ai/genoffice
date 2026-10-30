import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Direct w:ind on cell paragraphs: left/right/firstLine in twips, firstLineChars in
# em of the run size, and a bidi paragraph whose left indent sits at the right edge.
AR = 'عنوان التقرير'
rows = [
    para('indented left seven twenty, first line three sixty', ppr='<w:ind w:left="720" w:right="720" w:firstLine="360"/>'),
    para('two character first line indent', ppr='<w:ind w:firstLineChars="200"/>', run_rpr=rpr('Calibri', 24)),
    para(AR, ppr='<w:bidi/><w:ind w:left="720"/>', run_rpr=rpr('Calibri', 22, extra='<w:rtl/>')),
    para('control line'),
]
body = para('lead') + tbl(rows) + para('tail')
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
