import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
HAN = '文字格子'
def section(char_space, last):
    sect = A4 + f'<w:docGrid w:type="linesAndChars" w:linePitch="360" w:charSpace="{char_space}"/>'
    text = para(f'Latin abc 123 {HAN}', run_rpr=rpr('Calibri', 24, east_asia='SimSun'))
    if last:
        return text, sect
    return text + f'<w:p><w:pPr><w:sectPr>{sect}</w:sectPr></w:pPr></w:p>', None
p1, _ = section(11417, False)
p2, _ = section(8192, False)
p3, _ = section(4096, False)
p4, last = section(-820, True)
flags = "<w:useFELayout/>"
write(out_path(), p1 + p2 + p3 + p4, styles_xml(doc_defaults('Calibri', 24, east_asia='SimSun')), settings_xml(flags), sect=last)
