import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# docDefaults of the production paper: ascii Calibri, cs Arial, w:rtl, pPrDefault w:bidi
DEFAULTS = ('<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Arial"/>'
            '<w:sz w:val="22"/><w:szCs w:val="22"/><w:rtl/><w:lang w:val="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr><w:bidi/>' + spacing(after=160, line=259, rule='auto') + '</w:pPr></w:pPrDefault></w:docDefaults>')
AR = 'التعليم العالي: المفهوم والوظائف'
T = f'1.2 {AR} 2024'
PP = spacing(before=0, after=0, line=240, rule='auto')
def marker(s):
    return para(s, '<w:bidi w:val="0"/>' + spacing(before=120, after=0, line=240, rule='auto'),
                '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:rtl w:val="0"/></w:rPr>')
def lines(text, extra=''):
    return para([text] * 3, PP, f'<w:rPr>{extra}</w:rPr>' if extra else '')
ITEMS = [
    ('A1 no rPr', T, ''), ('A2 w:rtl', T, '<w:rtl/>'), ('A3 w:cs flag', T, '<w:cs/>'),
    ('A4 rFonts cs=Arial only', T, '<w:rFonts w:cs="Arial"/>'), ('A5 rtl val=0', T, '<w:rtl w:val="0"/>'),
    ('A6 cs+rtl', T, '<w:cs/><w:rtl/>'),
    ('A7 punct no digits', 'نص عربي. آخر: (ملاحظة) نسبة% وكذلك', ''),
    ('A8 arabic only', 'التعليم العالي المفهوم والوظائف', ''),
    ('A9 latin+digits no rPr', 'abc 1.2 def 2024', ''), ('A10 digits only no rPr', '1234567890 1234567890', ''),
    ('A11 digits w:rtl', '1234567890 1234567890', '<w:rtl/>'), ('A12 digits w:cs', '1234567890 1234567890', '<w:cs/>'),
    ('A13 TOC-like 1.6', '1.6 حدود الدراسة', ''),
    ('A14 arabic-indic digits no rPr', '١.٢ التعليم العالي ٢٠٢٤', ''),
]
body = ''.join(marker(lab) + lines(t, r) for lab, t, r in ITEMS)
TA = '<w:rFonts w:ascii="Traditional Arabic" w:hAnsi="Traditional Arabic" w:cs="Traditional Arabic"/><w:sz w:val="32"/><w:szCs w:val="32"/><w:rtl/>'
SA = '<w:rFonts w:ascii="Simplified Arabic" w:hAnsi="Simplified Arabic" w:cs="Simplified Arabic"/><w:sz w:val="28"/><w:szCs w:val="28"/><w:rtl/>'
S1 = 'هل واجهتي صعوبة في الحصول على مصادر مناسبة لتخصصك'
S2 = 'وقوع الحرج والمشقة الشديدة غير المعتادة'
for lab, t, r in [('B TA 16', S1, TA), ('B SA 14', S2, SA)]:
    body += marker(lab) + para(t, '<w:jc w:val="left"/>' + PP, f'<w:rPr>{r}</w:rPr>')
write(out_path(), body, styles_xml(DEFAULTS), settings_xml())
