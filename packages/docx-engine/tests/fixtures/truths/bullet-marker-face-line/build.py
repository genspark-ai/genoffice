import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Three-item bullet lists whose level rPr names Symbol (U+F0B7), Wingdings
# (U+F0A7) or no face (U+2022 in the text face), text and mark Calibri 10/11/12,
# single and line=276; a Normal declaring only an East Asian face (the text stays
# Calibri from docDefaults) mirrors the production template that hid the lift.
FACES = [('Symbol', '&#xF0B7;'), ('Wingdings', '&#xF0A7;'), (None, '&#x2022;')]
levels = []
for face, ch in FACES:
    r = f'<w:rPr><w:rFonts w:ascii="{face}" w:hAnsi="{face}" w:hint="default"/></w:rPr>' if face else ''
    levels.append([lvl('bullet', ch, '<w:ind w:left="360" w:hanging="360"/>', r)])
body = para('control 0')
k = 0
for num_id in (1, 2, 3):
    for sz in (20, 22, 24):
        for line in (240, 276):
            k += 1
            ppr = f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr>{spacing(after=0, line=line, rule="auto")}'
            rp = rpr(None, sz)
            for i in range(3):
                body += para(f'item {k}.{i} text line', ppr=ppr, run_rpr=rp)
            body += para(f'control {k}', run_rpr=rp)
normal_ea = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:eastAsia="DengXian"/></w:rPr></w:style>'
styles = f'{XML}<w:styles {W}>{doc_defaults("Calibri", 22)}{normal_ea}</w:styles>'
write(out_path(), body, styles, settings_xml(), numbering=numbering_xml(levels))
