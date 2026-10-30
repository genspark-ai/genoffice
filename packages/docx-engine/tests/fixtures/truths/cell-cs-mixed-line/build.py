import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# One cell line mixing a cs Arial Arabic run and a Calibri Latin run at one size:
# Word sizes the line by the tallest face (Calibri 1.22), the Arabic adds nothing.
AR = 'عنوان'
cs = ('<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Arial"/>'
      '<w:sz w:val="24"/><w:szCs w:val="24"/><w:rtl/></w:rPr>')
lat = rpr('Calibri', 24)
mixed = (f'<w:p><w:r>{lat}<w:t xml:space="preserve">Report </w:t></w:r><w:r>{cs}<w:t>{AR}</w:t></w:r>'
         f'<w:r>{lat}<w:t xml:space="preserve"> 2026</w:t></w:r></w:p>')
rows = [mixed, para(AR, ppr='<w:bidi/>', run_rpr=cs), para('Latin control', run_rpr=lat)]
body = para('lead', run_rpr=lat) + tbl(rows) + para('tail', run_rpr=lat)
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
