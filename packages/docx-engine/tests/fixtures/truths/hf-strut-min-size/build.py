import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _hfprobe import *
# Mixed run sizes in one strip paragraph: Word sizes each line by the runs on it,
# so the 8pt br-lines under an 11pt first line pitch 8 x 1.15, not 11 x 1.15.
# S1 11pt bold TNR <br/> 8pt TNR <br/> 8pt TNR; S2 one line of 8pt + 11pt runs;
# S3 four space-only sz 32 runs under sz 16 marks over a Calibri line (the mark
# sizes the line); S4 control: three 8pt lines; S5 control: three 11pt lines.
T8 = rp(rfonts('Times New Roman'), sz(16))
T11B = rp(rfonts('Times New Roman'), '<w:b/>', sz(22))
C11 = rp(rfonts('Calibri'), sz(22))
headers = [
    hp('', run('Title Line', T11B), br(T11B), run('second line small', T8), br(T8), run('third line small', T8)),
    hp('', run('Small ', T8), run('and large', T11B)),
    hp(rp(rfonts('Times New Roman'), sz(16)), run(' ', rp(rfonts('Times New Roman'), sz(32)))) * 4 + hp('', run('End line', C11)),
    hp('', run('first small', T8), br(T8), run('second small', T8), br(T8), run('third small', T8)),
    hp('', run('First Large', T11B), br(T11B), run('Second Large', T11B), br(T11B), run('Third Large', T11B)),
]
hf_doc(out_path(), headers)
