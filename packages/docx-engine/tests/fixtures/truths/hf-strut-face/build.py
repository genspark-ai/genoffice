import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _hfprobe import *
# Letterhead strips of Times New Roman 8pt under a Cambria 11 Normal: Word
# pitches 8 x 1.15 per line (the run face), the strip never takes the body face.
# S1 six TNR 8pt paragraphs; S2 six Calibri 8pt paragraphs; S3 one TNR 8pt line;
# S4 empty paragraph with a TNR sz 16 mark; S5 control Cambria 11 line.
T8 = rp(rfonts('Times New Roman'), sz(16))
C8 = rp(rfonts('Calibri'), sz(16))
SP = spacing(after=0, line=240, rule='auto')
lines = ['Department of Examples', 'Street 1, Example City', 'Phone 000 000 000', 'mail at example', 'NPSN : 2056', 'Reference 12']
headers = [
    ''.join(hp(SP, run(t, T8)) for t in lines),
    ''.join(hp(SP, run(t, C8)) for t in lines),
    hp(SP, run('NPSN : 2056', T8)),
    hp(SP + rp(rfonts('Times New Roman'), sz(16))),
    hp(SP, run('Cambria eleven control', rp(rfonts('Cambria'), sz(22)))),
]
hf_doc(out_path(), headers, defaults=doc_defaults('Cambria', 22, ppr=spacing(after=0, line=240, rule='auto')), body_rpr=rpr('Cambria', 22))
