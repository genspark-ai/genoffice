import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Letter, docDefaults Calibri 11. A missing Latin name ("Google Sans Text", fontTable altName
# Calibri, no PANOSE) at 11pt, line 240 and line 276, next to a name with a PANOSE-only entry
# (no altName) that Word substitutes by its generic rule.
SP = spacing(after=0, line=240, rule='auto')
SP276 = spacing(after=0, line=276, rule='auto')
FONTS = font_table_xml([
    ('Google Sans Text', {'altName': 'Calibri', 'charset': '00', 'family': 'auto', 'pitch': 'variable'}),
    ('Probe Serif Net', {'panose1': '02040503050406030204', 'charset': '00', 'family': 'roman', 'pitch': 'variable'}),
])
body = (
    para('Google Sans Text sample line', ppr=SP, run_rpr=rpr('Google Sans Text', 22))
    + para('Google Sans Text at 276', ppr=SP276, run_rpr=rpr('Google Sans Text', 22))
    + para('Probe Serif Net sample line', ppr=SP, run_rpr=rpr('Probe Serif Net', 22))
)
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml(), font_table=FONTS)
