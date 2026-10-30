import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _hfprobe import *
# Korean header text wraps at spaces only (Word's wordWrap default), never
# inside a word; S1 Malgun Gothic 11 long line in the 468pt column; S2 the same
# with w:wordWrap=0 (character wrapping allowed); S3 Latin control.
KO = ('한국어 머리글 줄바꿈 예제입니다 '
      '긴단어연결테스트를위한문장입니다 '
      '단어 사이에서만 줄이 바뀝니다 '
      '그리고 아주긴단어가하나더있습니다 '
      '마지막 단어')
M11 = rp(rfonts('Malgun Gothic', east_asia='Malgun Gothic'), sz(22))
headers = [
    hp('', run(KO, M11)),
    hp('<w:wordWrap w:val="0"/>', run(KO, M11)),
    hp('', run('Latin control line', rp(rfonts('Calibri'), sz(22)))),
]
hf_doc(out_path(), headers)
