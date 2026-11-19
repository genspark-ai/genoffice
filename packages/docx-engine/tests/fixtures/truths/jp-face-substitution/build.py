import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
faces = ['Meiryo UI', 'メイリオ', 'BIZ UDPゴシック', 'BIZ UDP明朝 Medium', 'MS Gothic', 'ＭＳ ゴシック', 'Aptos']
body = ''.join(para('あいう abc', run_rpr=rpr(f, 24, east_asia=f)) for f in faces)
write(out_path(), body, styles_xml(doc_defaults('Calibri', 24)), settings_xml())
