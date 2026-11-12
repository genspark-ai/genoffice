import zipfile, sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _table import *
OUT = sys.argv[1]
S = 'mmmm ' * 14 + 'mmmm'
W5 = 'mmmm ' * 4 + 'mmmm'
W9 = 'mmmm ' * 8 + 'mmmm'
G = [3009, 3009, 3008]
pct = lambda v: f'<w:tblW w:w="{v}" w:type="pct"/>'
tests = [
 ('P1 pct50 a a a', ['a', 'a', 'a'], 2500),
 ('P2 pct50 a mmmm.mmmm a', ['a', 'mmmm mmmm', 'a'], 2500),
 ('P3 pct50 a S a', ['a', S, 'a'], 2500),
 ('P4 pct50 mmmm S W5', ['mmmm', S, W5], 2500),
 ('P5 pct100 a S S', ['a', S, S], 5000),
 ('P6 pct100 S mmmm.mmmm W5', [S, 'mmmm mmmm', W5], 5000),
 ('P7 pct100 W5 S W9', [W5, S, W9], 5000),
 ('P8 pct100 a mmmm.mmmm a', ['a', 'mmmm mmmm', 'a'], 5000),
 ('P9 pct100 mm W5 W9', ['mm', W5, W9], 5000),
 ('P10 pct50 W5 S W9', [W5, S, W9], 2500),
 ('P11 pct100 8m.8m S a', ['mmmmmmmm mmmmmmmm', S, 'a'], 5000),
 ('P12 pct50 mm W5 W9', ['mm', W5, W9], 2500),
]
body = ''
for i, (name, texts, p) in enumerate(tests):
    rows = [['A', 'B', 'C'], texts]
    body += para(f'{i+1} {name}', brk=(i > 0 and i % 4 == 0)) + table(G, None, rows, tblw=pct(p)) + para('end')
doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}{sect}</w:body></w:document>'
with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', ct); z.writestr('_rels/.rels', rels)
    z.writestr('word/document.xml', doc); z.writestr('word/styles.xml', styles)
    z.writestr('word/settings.xml', settings); z.writestr('word/_rels/document.xml.rels', drels)
print(len(tests), 'tables')
