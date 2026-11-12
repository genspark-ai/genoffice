import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
H = rpr(sz=28)
HB = spacing(before=480, after=0)
HD_STYLE = ('<w:style w:type="paragraph" w:styleId="Hd"><w:name w:val="Hd"/><w:basedOn w:val="Normal"/>'
            f'<w:pPr>{HB}</w:pPr></w:style>')
def sect_para(after, text=None, typ=None):
    t = f'<w:type w:val="{typ}"/>' if typ else ''
    runs = f'<w:r>{H}<w:t>{text}</w:t></w:r>' if text else ''
    return f'<w:p><w:pPr>{spacing(after=after)}<w:sectPr>{t}{LETTER}</w:sectPr>{H}</w:pPr>{runs}</w:p>'
def h(text, ppr=HB): return para(text, ppr=ppr, run_rpr=H)
def f(text): return para(text, run_rpr=H)
body = ''.join([
    f('A1'), f('A2'), sect_para(200),            # p1 control; prev empty after=200
    h('B1'), f('B2'), sect_para(0),              # p2 expect +14; prev after=0
    h('C1'), f('C2'), sect_para(600),            # p3 expect +24; prev after=600
    h('D1'), f('D2'), sect_para(200, 'D3'),      # p4 expect 0; prev text after=200
    h('E1'), f('E2'), sect_para(200),            # p5 expect +14
    h('F1', '<w:pStyle w:val="Hd"/>'), f('F2'),  # p6 expect +14 (style before)
    sect_para(200, 'F3', 'continuous'), h('G1'), f('G2'),  # continuous: G1 gap after F3
])
mode = int(sys.argv[2]) if len(sys.argv) > 2 else 15
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22, ppr=spacing(after=200, line=240, rule='auto')), HD_STYLE), settings_xml('', mode))
