import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Letter, 1125-twip margins: column 9990 twips, default 720 grid, Arial 11 single lines.
# The text ends at ~2594 twips: eleven trailing tabs reach the 10080 stop past the edge, ten end on 9360
A = rpr('Arial', 22)
SECT = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1125" w:bottom="1440" w:left="1125" w:header="720" w:footer="720" w:gutter="0"/>'
SP = spacing(after=0, line=240, rule='auto')
TEXT = 'Trailing tabs past the edge'
body = (
    para(TEXT + '\t' * 11 + ' ' * 20, ppr=SP, run_rpr=A)
    + para(TEXT + '\t' * 10 + ' ' * 20, ppr=SP, run_rpr=A)
    + para('NEXT', ppr=SP, run_rpr=A)
    + para(TEXT + '\t' * 11 + 'end', ppr=SP, run_rpr=A)
    + para('LAST', ppr=SP, run_rpr=A)
    + para(TEXT + '\t' * 10 + 'end', ppr=SP, run_rpr=A)
    + para('TAIL', ppr=SP, run_rpr=A)
)
write(out_path(), body, styles_xml(doc_defaults('Arial', 22)), settings_xml(), sect=SECT)
