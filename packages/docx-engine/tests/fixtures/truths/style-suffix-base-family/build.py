import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Letter, docDefaults Calibri 11. Style-suffixed spellings of faces Word ships or
# the machine installs: Word drops the style word / MT tag and lays out the base
# family (word-truths 2026-10-22), not the Calibri/Cambria substitute.
SP = spacing(after=0, line=240, rule='auto')
FONTS = font_table_xml([
    ('Arial Bold', {'panose1': '020B0604020202090204', 'charset': '00', 'family': 'swiss', 'pitch': 'variable'}),
    ('Courier New Regular', {'panose1': '02070309020205020404', 'charset': '00', 'family': 'modern', 'pitch': 'fixed'}),
    ('Bookman Old Style Bold', {'panose1': '02050604050505020204', 'charset': '00', 'family': 'roman', 'pitch': 'variable'}),
    ('Monotype Corsiva Italic', {'panose1': '03010101010201010101', 'charset': '00', 'family': 'script', 'pitch': 'variable'}),
])
body = ''.join(
    para(f'{name} sample line', ppr=SP, run_rpr=rpr(name, 24))
    for name in ('Arial Bold', 'Courier New Regular', 'Bookman Old Style Bold', 'Monotype Corsiva Italic')
)
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml(), font_table=FONTS)
