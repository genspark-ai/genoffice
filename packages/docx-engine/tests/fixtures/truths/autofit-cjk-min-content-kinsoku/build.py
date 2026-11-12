import zipfile, sys, json, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _table import *
OUT = sys.argv[1]
M30 = 'm' * 30
G4 = [2256, 2257, 2256, 2257]
K = [('K0 a control', 'a'), ('K1 ai (no kinsoku)', 'あい'), ('K2 a + ideographic comma', 'あ、'), ('K3 open paren + a', '（あ'),
     ('K4 aiu, eo', 'あいう、えお'), ('K5 aaaa close paren', 'ああああ）'), ('K6 a prolonged i', 'あーい'),
     ('K7 a small-a i', 'あぁい'), ('K8 a SNS i', 'あSNSい'), ('K9 kanji x4', '広告宅配'), ('K10 a sok-on i', 'あっい'), ('K11 a middle dot i', 'あ・い')]
def run_ea(t, b=False):
    t = t.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
    rpr = '<w:rPr>' + ('<w:b/>' if b else '') + '<w:rFonts w:eastAsia="游明朝"/><w:lang w:eastAsia="ja-JP"/></w:rPr>'
    return f'<w:r>{rpr}<w:t xml:space="preserve">{t}</w:t></w:r>'
def cell_ea(tcw, text, b=False):
    return f'<w:tc><w:tcPr><w:tcW w:w="{tcw}" w:type="dxa"/></w:tcPr><w:p>{run_ea(text, b)}</w:p></w:tc>'
def table_ea(grid, rows):
    g = ''.join(f'<w:gridCol w:w="{w}"/>' for w in grid)
    trs = ''.join('<w:tr>' + ''.join(cell_ea(grid[ci], txt, ri == 0) for ci, txt in enumerate(row)) + '</w:tr>' for ri, row in enumerate(rows))
    return f'<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid>{g}</w:tblGrid>{trs}</w:tbl>'
body = ''
for i, (name, k) in enumerate(K):
    body += para(f'{i+1} {name}', brk=(i > 0 and i % 4 == 0)) + table_ea(G4, [['A', 'B', 'C', 'D'], [k, M30, M30, M30]]) + para('end')
doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}{sect}</w:body></w:document>'
with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', ct); z.writestr('_rels/.rels', rels)
    z.writestr('word/document.xml', doc); z.writestr('word/styles.xml', styles)
    z.writestr('word/settings.xml', settings); z.writestr('word/_rels/document.xml.rels', drels)
print(len(K), 'tables')
