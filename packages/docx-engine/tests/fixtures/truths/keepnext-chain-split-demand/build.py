import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
C = rpr('Calibri', 22)
SP = spacing(after=0, line=240, rule='auto')
filler = para([f'filler {i+1}' for i in range(20)], ppr=SP, run_rpr=C)
keep3 = ''.join(para([f'keep {k+1} line {j+1}' for j in range(3)], ppr='<w:keepNext/><w:widowControl w:val="0"/>' + SP, run_rpr=C) for k in range(60))
body = filler + keep3 + para('anchor', ppr=SP, run_rpr=C)
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml(), sect=A4)
