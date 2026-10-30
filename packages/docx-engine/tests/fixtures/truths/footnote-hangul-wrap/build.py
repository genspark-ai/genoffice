import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Footnote bodies: a Korean Malgun Gothic 10 note wraps at spaces only in the
# 468 pt column; a Latin control note (Calibri 10) and a note whose runs are
# 9 pt Times New Roman (the line pitch follows the runs' face, not the style).
# Body: three Calibri 11 paragraphs, each carrying one reference.
KO = ('한국어 각주 줄바꿈 예제입니다 '
      '긴단어연결테스트를위한문장입니다 '
      '단어 사이에서만 줄이 바퀌야 합니다 '
      '그리고 아주긴단어가하나더있습니다 '
      '마지막 단어입니다 끝')
M10 = '<w:rPr><w:rFonts w:ascii="Malgun Gothic" w:hAnsi="Malgun Gothic" w:eastAsia="Malgun Gothic"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
C10 = rpr('Calibri', 20)
T9 = rpr('Times New Roman', 18)
def note(text, r):
    ref = '<w:r>' + r.replace('</w:rPr>', '<w:vertAlign w:val="superscript"/></w:rPr>') + '<w:footnoteRef/></w:r>'
    return f'<w:p>{ref}<w:r>{r}<w:t xml:space="preserve"> {text}</w:t></w:r></w:p>'
LAT = 'Latin control footnote with enough words to wrap onto a second line inside the text column of the page for the probe measurement here.'
notes = [note(KO, M10), note(LAT, C10), note(LAT, T9)]
SP = spacing(after=0, line=240, rule='auto')
body = ''.join(
    f'<w:p><w:pPr>{SP}</w:pPr><w:r>{rpr("Calibri", 22)}<w:t xml:space="preserve">Body paragraph {i} with a note</w:t></w:r>{footnote_ref(i)}</w:p>'
    for i in (1, 2, 3)
)
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml(), footnotes=footnotes_xml(notes))
