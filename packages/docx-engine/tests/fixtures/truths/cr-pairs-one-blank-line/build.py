import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
C = rpr('Calibri', 24)
body = ('<w:p><w:r>' + C + '<w:t>alpha</w:t></w:r><w:r>' + C + '<w:cr/></w:r><w:r>' + C + '<w:cr/></w:r><w:r>' + C + '<w:t>bravo</w:t></w:r></w:p>'
        + '<w:p><w:r>' + C + '<w:t>alpha</w:t></w:r><w:r>' + C + '<w:br/></w:r><w:r>' + C + '<w:br/></w:r><w:r>' + C + '<w:t>bravo</w:t></w:r></w:p>'
        + para('Cambria Math title', run_rpr=rpr('Cambria Math', 40, extra='<w:b/>')))
write(out_path(), body, styles_xml(doc_defaults('Calibri', 24)), settings_xml())
