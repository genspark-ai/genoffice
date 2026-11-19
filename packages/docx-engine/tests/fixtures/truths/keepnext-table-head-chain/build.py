import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
KEEP = '<w:keepNext/>'
def row(label, header=False, keep_next=False, cant_split=True):
    trpr = ('<w:tblHeader/>' if header else '') + ('<w:cantSplit/>' if cant_split else '')
    ppr = KEEP if keep_next else ''
    cells = ''.join(f'<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>{para(f"{label} c{c}", ppr=ppr)}</w:tc>' for c in range(2))
    return f'<w:tr><w:trPr>{trpr}</w:trPr>{cells}</w:tr>'
rows = row('header', header=True) + ''.join(row(f'chain {i}', keep_next=True) for i in range(3)) + row('body 1') + row('body 2')
table = f'<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>{rows}</w:tbl>'
body = ''.join(para(f'filler {i}') for i in range(5)) + para('Heading kept with its table', ppr=KEEP) + table + para('after the table')
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
