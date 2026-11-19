import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _hfprobe import *
# Header lines whose runs declare no face: Word sizes them by their size x the
# inherited face's factor (docDefaults Calibri / eastAsia SimSun), like the body.
# Three br-lines per paragraph so the pitch is the line height; the empty-mark
# cases stack four empty paragraphs over a Calibri line so the push measures.
# S1 sz 32 inherited Calibri; S2 sz 16; S3 everything inherited (11pt); S4 CJK
# run inheriting SimSun sz 24; S5 four empty paragraphs, mark Times New Roman
# sz 16 (no hint); S6 four empty paragraphs, no mark rPr; S7 Calibri 11 control.
CJK = '文档标题'
C11 = rp(rfonts('Calibri'), sz(22))
def three(text, r):
    return hp('', run(text + ' one', r), br(r), run(text + ' two', r), br(r), run(text + ' three', r))
headers = [
    three('Sixteen inherited', rp(sz(32))),
    three('Eight inherited', rp(sz(16))),
    three('Inherits everything', ''),
    three(CJK, rp(sz(24))),
    hp(rp(rfonts('Times New Roman'), sz(16))) * 4 + hp('', run('End line', C11)),
    hp('') * 4 + hp('', run('End line', C11)),
    three('Calibri declared', C11),
]
hf_doc(out_path(), headers, defaults=doc_defaults('Calibri', 22, east_asia='SimSun'))
