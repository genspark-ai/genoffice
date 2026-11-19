import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
from _hfprobe import *
# Header paragraphs under an auto line multiple far below 1 (w:line 14 = 0.058)
# and a mild one (120 = 0.5): does Word paint the glyph whole over the tiny line
# box, and how far does the strip advance? Right-aligned Calibri 8 text under a
# sz 1 paragraph mark (HTML-converted letterhead), one nextPage section each.
T = rp(rfonts('Calibri'), sz(16))
line = lambda v: spacing(line=v, rule='auto') + '<w:jc w:val="right"/>' + rp(sz(2))
headers = [
    hp(line(14), run('Supply Agreement header', T)),
    hp(line(120), run('Supply Agreement header', T)),
    hp(line(240), run('Supply Agreement header', T)),
    hp(line(14), run('Supply Agreement header', T), br(T), run('second line', T)),
    hp(line(14), run('Supply Agreement header', T)) + hp(line(14), run('second paragraph', T)),
]
hf_doc(out_path(), headers)
