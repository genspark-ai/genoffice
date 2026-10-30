import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# body-only twin of render-grid-strut-two-face: 8.5pt / 10.5pt two-face CJK lines
# (ascii Calibri + eastAsia Microsoft YaHei) and a single-face Microsoft YaHei control
HAN = '文字段落行高测试'
SP = spacing(after=0, line=300, rule='auto')
body = (
    para(HAN, ppr=SP, run_rpr=rpr('Calibri', 17, east_asia='Microsoft YaHei'))
    + para(HAN, ppr=SP, run_rpr=rpr('Calibri', 21, east_asia='Microsoft YaHei'))
    + para(HAN, ppr=SP, run_rpr=rpr('Microsoft YaHei', 17, east_asia='Microsoft YaHei'))
    + para('Latin only', ppr=SP, run_rpr=rpr('Calibri', 17, east_asia='Microsoft YaHei'))
)
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22, east_asia='Microsoft YaHei', ppr=SP)), settings_xml(compat_mode=14))
