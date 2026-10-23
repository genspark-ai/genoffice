import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
C = rpr('Calibri', 22)
body = (para('约 6 万 BIM', ppr='<w:autoSpaceDE w:val="0"/>', run_rpr=C)
        + para('约 6 万 BIM', ppr='<w:autoSpaceDN w:val="0"/>', run_rpr=C)
        + para('约 6 万 BIM', ppr='<w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>', run_rpr=C)
        + para('约 6 万 BIM', run_rpr=C))
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml('<w:useFELayout/>', 14))
