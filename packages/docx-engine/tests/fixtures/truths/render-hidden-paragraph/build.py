import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
AR = rpr('Arial', 21)
V = rpr('Arial', 21, extra='<w:vanish/>')
body = (para('P0 after 120', ppr=spacing(after=120), run_rpr=AR)
        + para('hidden', ppr=spacing(before=120, after=120), run_rpr=V, mark_rpr=V)
        + para('P2 before 120', ppr=spacing(before=120), run_rpr=AR)
        + '<w:p><w:pPr>' + spacing(after=120) + '</w:pPr><w:bookmarkStart w:id="0" w:name="probe"/><w:r>' + AR + '<w:t>P3 after 120</w:t></w:r></w:p>'
        + '<w:bookmarkEnd w:id="0"/>'
        + para('P4 before 120', ppr=spacing(before=120), run_rpr=AR))
write(out_path(), body, styles_xml(doc_defaults('Arial', 21)), settings_xml(), sect=A4)
