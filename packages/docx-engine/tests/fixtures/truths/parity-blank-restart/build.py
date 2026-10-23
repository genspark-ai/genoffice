import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
def doc(even_odd):
    s1 = para('section one') + f'<w:p><w:pPr><w:sectPr>{LETTER}</w:sectPr></w:pPr></w:p>'
    s2 = para('section two restarts at 1')
    settings = f'{XML}<w:settings {W}>' + ('<w:evenAndOddHeaders/>' if even_odd else '') + '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'
    return s1 + s2, settings
body, st = doc(True)
write(out_path(), body, styles_xml(doc_defaults()), st, sect=LETTER + '<w:pgNumType w:start="1"/><w:type w:val="nextPage"/>')
body, st = doc(False)
write('probe-off.docx', body, styles_xml(doc_defaults()), st, sect=LETTER + '<w:pgNumType w:start="1"/><w:type w:val="nextPage"/>')
