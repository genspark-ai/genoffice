import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# A4, 1080-twip margins: column 9746 twips; TOC lines "title<tab>page" with a right
# dot-leader stop at the column edge under left indents 0 / 240 / 480
A = rpr('Arial', 22)
SECT = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" w:header="708" w:footer="708" w:gutter="0"/>'
SP = spacing(after=0, line=240, rule='auto')
TABS = '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9746"/></w:tabs>'


def entry(title, page, left):
    ind = f'<w:ind w:left="{left}"/>' if left else ''
    return para(f'{title}\t{page}', ppr=TABS + ind + SP, run_rpr=A)


body = (
    entry('Chapter one', '3', 0)
    + entry('Section one point one', '5', 240)
    + entry('Sub section one point one point one', '7', 480)
    + entry('Chapter two', '12', 0)
    + entry('Section two point one', '14', 240)
    + para('Control body paragraph', ppr=SP, run_rpr=A)
)
write(out_path(), body, styles_xml(doc_defaults('Arial', 22)), settings_xml(), sect=SECT)
