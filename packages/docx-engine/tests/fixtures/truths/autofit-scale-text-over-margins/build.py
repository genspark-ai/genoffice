import zipfile, sys, json, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _table import *
OUT = sys.argv[1]
m = lambda n: 'm' * n
G3 = [3009, 3009, 3008]; G4 = [2256, 2257, 2256, 2257]
tests = [
 ('S1 40m 20m 10m uniform', G3, dxa(G3), [m(40), m(20), m(10)]),
 ('S2 40m 5m 5m 5m uniform', G4, dxa(G4), [m(40), m(5), m(5), m(5)]),
 ('S3 same, tcW 1500/2500/2500/2526', G4, dxa([1500, 2500, 2500, 2526]), [m(40), m(5), m(5), m(5)]),
 ('S4 40m (5m 5m 5m) 5m 5m: max enters?', G4, dxa(G4), [m(40), f'{m(5)} {m(5)} {m(5)}', m(5), m(5)]),
 ('S5 60m a a', G3, dxa(G3), [m(60), 'a', 'a']),
 ('S6 14 cols corpus-like', [644] * 14, dxa([644] * 14), [m(k) for k in [14, 6, 6, 5, 6, 3, 5, 7, 5, 7, 5, 4, 7, 9]]),
 ('S7 40m 5m 5m 5m no tcW', G4, None, [m(40), m(5), m(5), m(5)]),
 ('S8 40m 20m 10m tblW dxa 9026', G3, dxa(G3), [m(40), m(20), m(10)]),
]
body = ''
for i, (name, grid, tcws, texts) in enumerate(tests):
    rows = [[chr(65 + k) for k in range(len(grid))], texts]
    tblw = '<w:tblW w:w="9026" w:type="dxa"/>' if 'dxa 9026' in name else '<w:tblW w:w="0" w:type="auto"/>'
    body += para(f'{i+1} {name}', brk=(i > 0 and i % 3 == 0)) + table(grid, tcws, rows, tblw=tblw) + para('end')
doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}{sect}</w:body></w:document>'
with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', ct); z.writestr('_rels/.rels', rels)
    z.writestr('word/document.xml', doc); z.writestr('word/styles.xml', styles)
    z.writestr('word/settings.xml', settings); z.writestr('word/_rels/document.xml.rels', drels)
print(len(tests), 'tables')
