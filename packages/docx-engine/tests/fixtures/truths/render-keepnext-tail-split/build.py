import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Letter 1440 margins = 648pt body = 54 exact-12pt lines, Arial 12, widow control on.
# Page 1: 43 filler lines, then keepNext A(4) B(3, keepLines) C(3) over a 2-line anchor:
# 11 lines of room, C cannot keep the anchor and cannot split, B is atomic, so A splits 2 + 2.
# Page 3: 48 filler lines, keepNext heading + 6-line keepNext body over a 2-line anchor:
# the chain would fit a fresh page, yet the body splits 4 + 2 at the page bottom.
A = rpr('Arial', 24)
SP = spacing(after=0, line=240, rule='exact')
WC = '<w:widowControl/>'
KN = '<w:keepNext/>'
KL = '<w:keepLines/>'


def lines(tag, n, ppr):
    return para([f'{tag} line {i + 1}' for i in range(n)], ppr=ppr, run_rpr=A)


body = (
    lines('filler', 43, WC + SP)
    + lines('A', 4, KN + WC + SP)
    + lines('B', 3, KN + KL + WC + SP)
    + lines('C', 3, KN + WC + SP)
    + lines('anchor', 2, WC + SP)
    + lines('filler2', 48, '<w:pageBreakBefore/>' + WC + SP)
    + lines('heading', 1, KN + KL + WC + SP)
    + lines('body', 6, KN + WC + SP)
    + lines('anchor2', 2, WC + SP)
)
write(out_path(), body, styles_xml(doc_defaults('Arial', 24)), settings_xml())
