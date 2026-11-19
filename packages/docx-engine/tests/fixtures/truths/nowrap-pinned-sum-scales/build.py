import sys
import os
BASE=os.path.join(os.path.dirname(os.path.abspath(__file__)), 'build_base.py')
sys.argv=[sys.argv[0], sys.argv[1]]
exec(open(BASE).read().split("long='")[0])
long='Bendahara penerima menerima tagihan dari pajak PBB sebesar Rp 11.000.000 untuk tahun berjalan dan mencatatnya'
body=''
r=lambda: [cell(1530,'02 Jan 2025',True,2),cell(12915,long)]
# P15: 056 minimal: + a row where both cells are noWrap (empty text)
body+=table('P15 span noWrap + row both noWrap empty', [926,49,8051], [r(),[cell(1530,' ',True,2),cell(12915,' ',True)]])
# P16: + row where only col2 is noWrap
body+=table('P16 span noWrap + row col2 noWrap only', [926,49,8051], [r(),[cell(1530,' ',False,2),cell(12915,' ',True)]])
# P17: both noWrap but the preferred sum fits (1530 + 7000 = 8530 < 9026)
body+=table('P17 both noWrap sum fits', [1530,7496], [[cell(1530,'02 Jan 2025',True),cell(7000,long,True)],[cell(1530,'03 Jan 2025',True),cell(7000,long,True)]])
# P18: both noWrap, sum 1530+9000=10530 slightly over
body+=table('P18 both noWrap sum slightly over', [1530,7496], [[cell(1530,'02 Jan 2025',True),cell(9000,long,True)],[cell(1530,'03 Jan 2025',True),cell(9000,long,True)]])
# P19: both noWrap with text in both (P15 w/ text)
body+=table('P19 both noWrap with text', [926,49,8051], [[cell(1530,'02 Jan 2025',True,2),cell(12915,long,True)]])
doc=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>'
src=open(BASE).read()
exec(src[src.index("styles=f'"):])
