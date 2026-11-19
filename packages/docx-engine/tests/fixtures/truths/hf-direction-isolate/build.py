import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _hfprobe import *
# Paragraph direction in a header strip: a w:bidi Arabic paragraph ends at the
# right margin and starts its line there; an LTR paragraph mixing digits and
# Latin keeps reading order from the left margin; a w:bidi paragraph holding
# only an LTR run (no w:rtl) still lays its text from the right margin.
AR = rp(rfonts('Calibri', cs='Arial'), '<w:rtl/>', sz(24))
L = rp(rfonts('Calibri'), sz(24))
headers = [
    hp('<w:bidi/>', run('مرحبا بكم في الرأس', AR)),
    hp('', run('abc 123 def', L)),
    hp('<w:bidi/>', run('abc 123 def', L)),
    hp('', run('العنوان ', AR), run('and Latin 123', L)),
]
hf_doc(out_path(), headers)
