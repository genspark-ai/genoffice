import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *

GLYPHS = ['◆', '■', '▪', '□', '▲', '▼', '▶', '►', '●', '○',
          '◇', '★', '☆', '✔', '✓', '→', '⚠', '✅',
          '▶️', '⚠️', '✔️', '◆️']


def label(g):
    return '+'.join(hex(ord(c)) for c in g)


def mixed(text_rpr, sym, line, sym_rpr=None):
    sr = sym_rpr if sym_rpr is not None else text_rpr
    parts = []
    for i in range(4):
        if i:
            parts.append(f'<w:r>{text_rpr}<w:br/></w:r>')
        parts.append(f'<w:r>{text_rpr}<w:t xml:space="preserve">Item line {i} </w:t></w:r>')
        parts.append(f'<w:r>{sr}<w:t xml:space="preserve">{sym}</w:t></w:r>')
        parts.append(f'<w:r>{text_rpr}<w:t xml:space="preserve"> tail text</w:t></w:r>')
    pp = f'<w:pPr>{spacing(before=0, after=0, line=line, rule="auto")}{text_rpr}</w:pPr>'
    return f'<w:p>{pp}{"".join(parts)}</w:p>'


def cjk(ea, sym, line, hint=False, sz=22):
    h = ' w:hint="eastAsia"' if hint else ''
    tr = f'<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="{ea}" w:cs="Calibri"{h}/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/><w:lang w:eastAsia="zh-CN"/></w:rPr>'
    parts = []
    for i in range(4):
        if i:
            parts.append(f'<w:r>{tr}<w:br/></w:r>')
        parts.append(f'<w:r>{tr}<w:t xml:space="preserve">中文测试行 {i} {sym} 符号文本</w:t></w:r>')
    pp = f'<w:pPr>{spacing(before=0, after=0, line=line, rule="auto")}{tr}</w:pPr>'
    return f'<w:p>{pp}{"".join(parts)}</w:p>'


def marker(text):
    return para(f'## {text}', run_rpr=rpr('Arial', 16))


body = ''
for tf, sz in (('Calibri', 22), ('Times New Roman', 24), ('Arial', 22)):
    tr = rpr(tf, sz)
    for line in (240, 276):
        body += marker(f'{tf} {sz // 2} plain line{line}')
        body += mixed(tr, 'o', line)
    for g in GLYPHS:
        for line in (240, 276):
            body += marker(f'{tf} {sz // 2} {label(g)} line{line}')
            body += mixed(tr, g, line)
# the corpus shape: Calibri body with a missing eastAsia face declared on the run
noto = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Noto Sans CJK SC" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'
for g in ('◆', '▲', '▶', '→', '★'):
    body += marker(f'Calibri 11 eastAsia NotoSansCJKSC {label(g)} line259')
    body += mixed(noto, g, 259)
for ea, hint in (('SimSun', False), ('SimSun', True), ('MS Mincho', True), ('Malgun Gothic', True), ('Noto Sans CJK SC', True)):
    for g in ('◆', '●', '▲', '★', '→', '■'):
        body += marker(f'CJK {ea} hint={hint} {label(g)} line240')
        body += cjk(ea, g, 240, hint)
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
