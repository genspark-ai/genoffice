import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
START = int(os.environ.get('PGNUM_START', '1'))
cover_sect = f'<w:sectPr>{LETTER}</w:sectPr>'
body = para('COVER') + f'<w:p><w:pPr>{cover_sect}</w:pPr></w:p>' + para('TOC')
tail = f'<w:sectPr><w:pgNumType w:fmt="lowerRoman" w:start="{START}"/>{LETTER}</w:sectPr>'
write(out_path(), body, styles_xml(doc_defaults('Arial', 24)), settings_xml(compat_flags='') .replace('<w:compat>', '<w:evenAndOddHeaders/><w:compat>'), sect_pr=tail)
