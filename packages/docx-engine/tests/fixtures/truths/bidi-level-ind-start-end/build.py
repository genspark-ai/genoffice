import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
A = rpr('Arial', 21)
AR = 'نص عربي طويل للتجربة'
levels = [
    [lvl('bullet', '•', '<w:ind w:right="360" w:hanging="360"/>', extra='<w:bidi/>')],
    [lvl('bullet', '•', '<w:ind w:left="360" w:hanging="360"/>', extra='<w:bidi/>')],
    [lvl('bullet', '•', '<w:ind w:hanging="360"/>', extra='<w:bidi/>')],
    [lvl('bullet', '•', '<w:ind w:left="720" w:right="360" w:hanging="360"/>', extra='<w:bidi/>')],
]
def item(num):
    return f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num}"/></w:numPr><w:jc w:val="both"/></w:pPr><w:r>{A}<w:rtl/><w:t>{AR}</w:t></w:r></w:p>'
body = ''.join(item(i + 1) for i in range(4))
sect = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1200" w:right="1200" w:bottom="1200" w:left="1200" w:header="708" w:footer="708" w:gutter="0"/>'
write(out_path(), body, styles_xml(doc_defaults('Arial', 21)), settings_xml(), sect=sect, numbering=numbering_xml(levels))
