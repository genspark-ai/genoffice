import zipfile, sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _table import *
OUT = sys.argv[1]
S = 'mmmm ' * 14 + 'mmmm'           # wrapping sentence: small min, huge max
T15 = 'm' * 15                      # ~132pt token
T30 = 'm' * 30                      # ~264pt token
T20 = 'm' * 20
# (name, grid/tcW per col, body row texts, tblW override or None)
U = [2256] * 4
N = [1800] * 4
tests = [
 ('R1 a a S T30', U, ['a', 'a', S, T30], None),
 ('R2 a a a T30 control', U, ['a', 'a', 'a', T30], None),
 ('R3 a mmmm.mmmm a T30', U, ['a', 'mmmm mmmm', 'a', T30], None),
 ('R4 a 4xmmmm a T30', U, ['a', 'mmmm mmmm mmmm mmmm', 'a', T30], None),
 ('R5 a 8m.8m a T30', U, ['a', 'mmmmmmmm mmmmmmmm', 'a', T30], None),
 ('R6 a a S T15', U, ['a', 'a', S, T15], None),
 ('R7 6m 6m.6m S T30', U, ['mmmmmm', 'mmmmmm mmmmmm', S, T30], None),
 ('R8 a mm mmmm T30', U, ['a', 'mm', 'mmmm', T30], None),
 ('R9 a a S T20', U, ['a', 'a', S, T20], None),
 ('R10 a S S T30', U, ['a', S, S, T30], None),
 ('U1 1800x4 a a S T15 under column', N, ['a', 'a', S, T15], None),
 ('U2 1800x4 a a S T30 over column', N, ['a', 'a', S, T30], None),
 ('U3 1800x4 a a a T20', N, ['a', 'a', 'a', T20], None),
 ('U4 1800x4 a a a T30', N, ['a', 'a', 'a', T30], None),
]
body = ''
for i, (name, grid, texts, tblw) in enumerate(tests):
    rows = [['A', 'B', 'C', 'D'], texts]
    kw = {'tblw': tblw} if tblw else {}
    body += para(f'{i+1} {name}', brk=(i > 0 and i % 4 == 0)) + table(grid, dxa(grid), rows, **kw) + para('end')
doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}{sect}</w:body></w:document>'
with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', ct); z.writestr('_rels/.rels', rels)
    z.writestr('word/document.xml', doc); z.writestr('word/styles.xml', styles)
    z.writestr('word/settings.xml', settings); z.writestr('word/_rels/document.xml.rels', drels)
print(len(tests), 'tables')
