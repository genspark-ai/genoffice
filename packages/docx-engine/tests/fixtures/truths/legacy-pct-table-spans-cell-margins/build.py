import zipfile, sys, json, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _table import *
OUT = sys.argv[1]; COMPAT = sys.argv[2] if len(sys.argv) > 2 else '14'
G = [3009, 3009, 3008]
pct = lambda v, extra='': f'<w:tblW w:w="{v}" w:type="pct"/>{extra}'
tests = [
 ('L1 pct5000 left', 5000, ''),
 ('L2 pct2500 left', 2500, ''),
 ('L3 pct5000 tblInd 300', 5000, '<w:tblInd w:w="300" w:type="dxa"/>'),
 ('L4 pct5000 jc center', 5000, '<w:jc w:val="center"/>'),
 ('L5 pct5000 cellMar 200', 5000, '<w:tblCellMar><w:left w:w="200" w:type="dxa"/><w:right w:w="200" w:type="dxa"/></w:tblCellMar>'),
 ('L6 dxa 9026 left', None, '<w:tblW w:w="9026" w:type="dxa"/>'),
]
body = ''
for i, (name, p, extra) in enumerate(tests):
    rows = [['A', 'B', 'C'], ['a', 'a', 'a']]
    tblw = pct(p, extra) if p else extra
    body += para(f'{i+1} {name}', brk=(i > 0 and i % 3 == 0)) + table(G, None, rows, tblw=tblw) + para('end')
doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}{sect}</w:body></w:document>'
st = settings.replace('w:val="15"', f'w:val="{COMPAT}"')
with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', ct); z.writestr('_rels/.rels', rels)
    z.writestr('word/document.xml', doc); z.writestr('word/styles.xml', styles)
    z.writestr('word/settings.xml', st); z.writestr('word/_rels/document.xml.rels', drels)
print(len(tests), 'tables compat', COMPAT)
