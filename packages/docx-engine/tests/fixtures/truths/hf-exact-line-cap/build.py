import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _hfprobe import *
# Fixed line rules in a header strip: exact 200 clips 11pt glyphs to a 10pt
# line, exact 480 with two br-lines is 48pt, atLeast 400 is 20pt per line,
# atLeast 200 under an 11pt run stays the natural 13.42 line; S5 control.
C11 = rp(rfonts('Calibri'), sz(22))
headers = [
    hp(spacing(line=200, rule='exact'), run('Exact ten point', C11)),
    hp(spacing(line=480, rule='exact'), run('Exact twenty-four', C11), br(C11), run('second line', C11)),
    hp(spacing(line=400, rule='atLeast'), run('At least twenty', C11), br(C11), run('second line', C11)),
    hp(spacing(line=200, rule='atLeast'), run('At least ten natural', C11)),
    hp('', run('Control line', C11)),
]
hf_doc(out_path(), headers)
