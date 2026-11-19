import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
faces = ['HG教科書体', 'HGP明朝B', 'HG行書体', 'ヒラギノ角ゴ Pro W3', 'HG丸ｺﾞｼｯｸM-PRO', 'ＭＳ Ｐ明朝', 'MS UI Gothic', 'Yu Gothic UI', 'UD デジタル 教科書体 N-R', '小塚明朝 Pr6N R', '架空宋体', '游明朝']
body = ''.join(para('あいう abc', ppr=spacing(after=0, line=240, rule='auto'), run_rpr=rpr(f, 20, east_asia=f)) for f in faces)
write(out_path(), body, styles_xml(doc_defaults('Calibri', 20)), settings_xml())
