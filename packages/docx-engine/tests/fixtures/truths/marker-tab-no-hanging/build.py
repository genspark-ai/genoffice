import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
C = rpr('Calibri', 22)
levels = [[lvl(ind='<w:ind w:left="1032" w:hanging="360"/>', extra='<w:tabs><w:tab w:val="num" w:pos="1032"/></w:tabs>')]]
def item(text, ppr=''):
    return f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>{ppr}</w:pPr><w:r>{C}<w:t>{text}</w:t></w:r></w:p>'
body = item('level indent') + item('para left 720 firstLine 0', '<w:ind w:left="720" w:firstLine="0"/>') + item('para left 200 firstLine 0', '<w:ind w:left="200" w:firstLine="0"/>')
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml(), numbering=numbering_xml(levels))
