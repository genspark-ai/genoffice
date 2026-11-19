import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *

BOX, BOXX = '☐', '☑'
GRIN, CHECK, CROSS = '\U0001F600', '✅', '❌'


def mixed(text_font, sz, sym_font, sym, line):
    tr = rpr(text_font, sz)
    sr = rpr(sym_font, sz) if sym_font else tr
    parts = []
    for i in range(4):
        if i:
            parts.append(f'<w:r>{tr}<w:br/></w:r>')
        parts.append(f'<w:r>{tr}<w:t xml:space="preserve">Item line {i} </w:t></w:r>')
        parts.append(f'<w:r>{sr}<w:t xml:space="preserve">{sym}</w:t></w:r>')
        parts.append(f'<w:r>{tr}<w:t xml:space="preserve"> tail text</w:t></w:r>')
    pp = f'<w:pPr>{spacing(before=0, after=0, line=line, rule="auto")}{tr}</w:pPr>'
    return f'<w:p>{pp}{"".join(parts)}</w:p>'


def bidi(line):
    ar = 'الأهلية: '
    tr = rpr('Times New Roman', 24)
    sr = rpr('Segoe UI Symbol', 24)
    parts = []
    for i in range(4):
        if i:
            parts.append(f'<w:r>{tr}<w:br/></w:r>')
        parts.append(f'<w:r>{tr}<w:t xml:space="preserve">{ar}</w:t></w:r><w:r>{sr}<w:t>{BOX}</w:t></w:r>'
                     f'<w:r>{tr}<w:t xml:space="preserve"> {ar}</w:t></w:r>')
    return f'<w:p><w:pPr><w:bidi/>{spacing(before=0, after=0, line=line, rule="auto")}{tr}</w:pPr>{"".join(parts)}</w:p>'


CASES = [(None, 'o'), ('Segoe UI Symbol', BOX), ('Segoe UI Symbol', BOXX), ('Segoe UI Symbol', 'o'), (None, BOX),
         ('Segoe UI Emoji', GRIN), ('Segoe UI Emoji', CHECK), (None, GRIN), (None, CHECK), (None, CROSS),
         ('Arial Unicode MS', CROSS)]
body = ''
for tf, sz in (('Times New Roman', 24), ('Calibri', 22), ('Arial', 22)):
    for sym_font, sym in CASES:
        for line in (240, 276):
            body += para(f'## {tf} {sz // 2} {sym_font} {hex(ord(sym))} line{line}', run_rpr=rpr('Arial', 16))
            body += mixed(tf, sz, sym_font, sym, line)
for line in (240, 276):
    body += para(f'## TNR 12 bidi Segoe UI Symbol {hex(ord(BOX))} line{line}', run_rpr=rpr('Arial', 16))
    body += bidi(line)
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
