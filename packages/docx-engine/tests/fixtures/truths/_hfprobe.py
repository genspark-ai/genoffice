"""Header-strip probe documents: one nextPage section per header part, a
Calibri body line per section, so each page's body top measures its strip."""
from _docx import *

BODY_SP = spacing(after=0, line=240, rule='auto')


def hf_doc(out, headers, defaults=None, extra_styles='', body_rpr=None, settings=None, page=LETTER):
    defaults = defaults if defaults is not None else doc_defaults('Calibri', 22)
    body_rpr = body_rpr if body_rpr is not None else rpr('Calibri', 22)
    settings = settings if settings is not None else settings_xml()
    hf = {}
    body = ''
    last = ''
    for i, h in enumerate(headers, 1):
        hf[f'header{i}.xml'] = hdr_xml(h)
        sect = f'<w:sectPr><w:headerReference w:type="default" r:id="rIdheader{i}"/>{page}</w:sectPr>'
        if i < len(headers):
            body += para(f'Section {i} body', ppr=BODY_SP + sect, run_rpr=body_rpr, mark_rpr='')
        else:
            last = sect
            body += para(f'Section {i} body', ppr=BODY_SP, run_rpr=body_rpr)
    write(out, body, styles_xml(defaults, extra_styles), settings, sect_pr=last, hf=hf)


def run(text, rpr=''):
    return f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'


def br(rpr=''):
    return f'<w:r>{rpr}<w:br/></w:r>'


def hp(ppr, *runs):
    return f'<w:p>{"<w:pPr>" + ppr + "</w:pPr>" if ppr else ""}{"".join(runs)}</w:p>'


def rfonts(ascii=None, east_asia=None, cs=None, hint=None):
    a = ''
    if ascii:
        a += f' w:ascii="{ascii}" w:hAnsi="{ascii}"'
    if east_asia:
        a += f' w:eastAsia="{east_asia}"'
    if cs:
        a += f' w:cs="{cs}"'
    if hint:
        a += f' w:hint="{hint}"'
    return f'<w:rFonts{a}/>' if a else ''


def rp(*kids):
    body = ''.join(kids)
    return f'<w:rPr>{body}</w:rPr>' if body else ''


def sz(v):
    return f'<w:sz w:val="{v}"/><w:szCs w:val="{v}"/>'
