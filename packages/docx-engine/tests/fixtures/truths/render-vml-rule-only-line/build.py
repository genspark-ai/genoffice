import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Paragraphs holding only a VML horizontal rule (<v:rect o:hr o:hrstd>, the <hr>
# of an HTML-origin document) under different run sizes, rule heights and line
# multiples; a Calibri 11 control line after each so the Word PDF gives the
# paragraph height. Case 2 is a title-page triple: rule / Georgia 14 bold text / rule.
VML = 'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"'


def hr(h=1.2, std=True, w=324):
    std_attr = ' o:hrstd="t"' if std else ''
    return (
        f'<w:pict><v:rect {VML} id="r" style="width:{w}pt;height:{h}pt" o:hralign="center"{std_attr}'
        ' o:hrnoshade="t" o:hr="t" fillcolor="#000000" stroked="f"/></w:pict>'
    )


G = rpr('Georgia', 28, extra='<w:b/>')
C = rpr('Calibri', 22)
L276 = spacing(after=0, line=276, rule='auto') + '<w:jc w:val="center"/>'
L240 = spacing(after=0, line=240, rule='auto') + '<w:jc w:val="center"/>'
ctrl = lambda k: para(f'control {k}')
rule_para = lambda ppr, run_rpr, h=1.2, std=True, mark=None: (
    f'<w:p><w:pPr>{ppr}{mark if mark is not None else run_rpr}</w:pPr><w:r>{run_rpr}{hr(h, std)}</w:r></w:p>'
)
body = ctrl(0)
body += rule_para(L276, G) + ctrl(1)
body += rule_para(L276, G) + para('Identifying Correcting Balancing', ppr=L276, run_rpr=rpr('Georgia', 28)) + rule_para(L276, G) + ctrl(2)
body += rule_para(L240, '') + ctrl(3)
body += rule_para(L240, G) + ctrl(4)
body += rule_para(L240, '', h=6) + ctrl(5)
body += rule_para(L240, '', h=20) + ctrl(6)
body += rule_para(L276, G, std=False) + ctrl(7)
body += rule_para(L240, '', mark=G) + ctrl(8)
body += rule_para(spacing(after=160, line=259, rule='auto'), '') + ctrl(9)
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml())
