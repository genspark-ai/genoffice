import sys; import os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
JA = ['日本語の段落テキスト行の高さを測る', '二行目の日本語テキストです', '三行目の日本語テキストです', '四行目の日本語テキストです', '五行目の日本語テキストです', '六行目の日本語テキストです']
for name, ea in [('yumincho', 'Yu Mincho'), ('msmincho', 'MS Mincho'), ('simsun', 'SimSun')]:
    box = textbox(para(JA), cx=4572000, cy=2000000, id=1)
    body = para(['Calibri lead line', 'second']) + para(JA) + para('mid') + box + para('tail')
    write(f'{name}.docx', body, styles_xml(doc_defaults('Calibri', 22, east_asia=ea)), settings_xml())
print('built')
