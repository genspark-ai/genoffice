import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# a missing complex-script face whose fontTable altName is a measured face (Sakkal Majalla),
# next to the same entry without altName, no entry at all, and the altName face itself
AR = ('أذاكر حاليا في مكتب مخصص داخل غرفتي الخاصة وهو مكان يتميز بالهدوء '
      'والإضاءة الجيدة إلا أنني أعاني من وجود هاتفي المحمول بجانبي بصورة مستمرة '
      'مما يسبب لي تشتتا ذهنيا كبيرا بسبب توالي الإشعارات والتنبيهات 11:00 ')
LAT = 'Hamburgefonstiv quick brown foxes jump over the lazy dog 0123456789 '
JA = '東京都港区芝公園の夢と希望 '
def ar(label, cs, run_sz=True, mark_cs=None):
    fonts = f'<w:rFonts w:ascii="Traditional Arabic" w:hAnsi="Traditional Arabic"{(" w:cs=%s" % chr(34)) + cs + chr(34) if cs else ""}/>'
    size = '<w:sz w:val="24"/><w:szCs w:val="24"/>' if run_sz else ''
    rp = f'<w:rPr>{fonts}{size}<w:rtl/></w:rPr>'
    mark = f'<w:rPr><w:rFonts w:ascii="{mark_cs}" w:hAnsi="{mark_cs}" w:cs="{mark_cs}"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>' if mark_cs else ''
    return f'<w:p><w:pPr><w:bidi/>{mark}</w:pPr><w:r>{rp}<w:t xml:space="preserve">{label} {AR * 3}</w:t></w:r></w:p>'
def lat(label, font):
    return para(f'{label} {LAT * 3}', run_rpr=rpr(font, 24))
def ja(label, ea):
    return para(f'{label} {JA * 6}', run_rpr=rpr('Calibri', 24, east_asia=ea))
body = (ar('P1', 'ProbeAltMajalla')
        + ar('P2', 'ProbeNoAlt')
        + ar('P3', 'ProbeNoEntry')
        + ar('P4', None, run_sz=False, mark_cs='ProbeAltMajalla')
        + ar('P5', 'Sakkal Majalla')
        + ar('P6', 'ProbeAltArial')
        + lat('P7', 'ProbeAltMajalla')
        + lat('P8', 'ProbeAltTimesSans')
        + lat('P9', 'ProbeNoAlt')
        + ja('P10', 'ProbeAltMincho')
        + ja('P11', 'ProbeNoAltEA'))
sig = {'charset': '00', 'family': 'roman', 'pitch': 'variable'}
ft = font_table_xml([
    ('ProbeAltMajalla', {'altName': 'Sakkal Majalla', **sig}),
    ('ProbeNoAlt', sig),
    ('ProbeAltArial', {'altName': 'Arial', **sig}),
    ('ProbeAltTimesSans', {'altName': 'Times New Roman', 'panose1': '020B0604020202020204', 'charset': '00', 'family': 'swiss', 'pitch': 'variable'}),
    ('ProbeAltMincho', {'altName': 'MS Mincho', 'charset': '80', 'family': 'roman', 'pitch': 'fixed'}),
    ('ProbeNoAltEA', {'charset': '80', 'family': 'roman', 'pitch': 'fixed'}),
    ('Traditional Arabic', {'panose1': '02020603050405020304', 'charset': 'B2', 'family': 'roman', 'pitch': 'variable'}),
    ('Calibri', {'panose1': '020F0502020204030204', 'charset': '00', 'family': 'swiss', 'pitch': 'variable'}),
])
dd = ('<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
      '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
      f'<w:pPrDefault><w:pPr>{spacing(after=160, line=259, rule="auto")}</w:pPr></w:pPrDefault></w:docDefaults>')
write(out_path(), body, styles_xml(dd), settings_xml(), font_table=ft, sect=A4)
