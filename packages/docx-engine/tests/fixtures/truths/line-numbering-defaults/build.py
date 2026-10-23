import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
def sec(ln):
    return f'<w:p><w:pPr><w:sectPr>{LETTER}{ln}</w:sectPr></w:pPr></w:p>'
body = (para('s1') + sec('<w:lnNumType w:countBy="1"/>')
        + para('s2') + sec('<w:lnNumType w:start="5"/>')
        + para('s3') + sec('<w:lnNumType w:start="0" w:restart="continuous"/>'))
body += para('s4')
write(out_path(), body, styles_xml(doc_defaults()), settings_xml(), sect=LETTER + '<w:lnNumType w:restart="newSection"/>')
