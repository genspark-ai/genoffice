import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
H = rpr(sz=28)
SB = spacing(before=400, after=200)
PBB = '<w:pageBreakBefore/>'
PBB_STYLE = ('<w:style w:type="paragraph" w:styleId="PbbHead"><w:name w:val="Pbb Head"/><w:basedOn w:val="Normal"/>'
             '<w:pPr>' + PBB + '</w:pPr></w:style>')
def p(text, ppr=''):
    return para(text, ppr=ppr, run_rpr=H)
body = ''.join([
    p('A1'),
    p('B1', PBB + SB),
    p('B2'),
    f'<w:p><w:pPr>{SB}{H}</w:pPr><w:r>{H}<w:br w:type="page"/></w:r><w:r>{H}<w:t>C1</w:t></w:r></w:p>',
    p('C2'),
    p('D1'),
    para(None, run_rpr=H).replace('</w:p>', f'<w:r>{H}<w:br w:type="page"/></w:r></w:p>'),
    p('E1', SB),
    p('E2'),
    p('F1', spacing(after=400)),
    p('G1', PBB + SB),
    p('G2'),
    p('H1', '<w:pStyle w:val="PbbHead"/>' + SB),
    p('H2'),
])
compat = sys.argv[2] if len(sys.argv) > 2 else ''
mode = int(sys.argv[3]) if len(sys.argv) > 3 else 15
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22), PBB_STYLE), settings_xml(compat, mode))
