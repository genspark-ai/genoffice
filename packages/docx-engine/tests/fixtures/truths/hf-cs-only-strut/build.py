import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _hfprobe import *
# Arabic-only header paragraphs lay out in their w:cs face alone: cs Arial 12
# paces 1.15 x 12, cs Sakkal Majalla 14 paces its own factor; the ascii Calibri
# slot contributes nothing. Three br-lines per paragraph so the pitch is the
# line. S3 Calibri 12 Latin control; S4 mixed Calibri Latin + cs Arial Arabic
# lines (the Calibri line).
AR = 'مرحبا بكم في الرأس'
A12 = rp(rfonts('Calibri', cs='Arial'), '<w:rtl/>', sz(24))
S14 = rp(rfonts('Calibri', cs='Sakkal Majalla'), '<w:rtl/>', sz(28))
C12 = rp(rfonts('Calibri'), sz(24))
def three(r, text=AR):
    return hp('<w:bidi/>', run(text, r), br(r), run(text, r), br(r), run(text, r))
mixed = run('Latin ', C12) + run(AR, A12) + run(' tail', C12)
headers = [
    three(A12),
    three(S14),
    hp('', run('Latin control twelve', C12), br(C12), run('Latin control twelve', C12), br(C12), run('Latin control twelve', C12)),
    hp('', mixed, br(C12), mixed, br(C12), mixed),
]
hf_doc(out_path(), headers)
