import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
T = rpr('Times New Roman', 24)
def spacer(tag, exact, pbb=False):
    extra = '<w:pageBreakBefore/>' if pbb else ''
    return para(f'spacer {tag}', ppr=extra + spacing(after=0, line=exact, rule='exact'), run_rpr=T)
def filler(n, after):
    return para([f'line {i+1}' for i in range(n)], ppr=spacing(after=after, line=240, rule='auto'), run_rpr=T)
body = (spacer('S1', 220) + filler(45, 120) + para('MARK S1', ppr=spacing(after=0, line=240, rule='auto'), run_rpr=T)
        + spacer('S11', 220, True) + filler(45, 0) + para('MARK S11', ppr=spacing(after=0, line=240, rule='auto'), run_rpr=T)
        + spacer('T1', 100, True) + filler(45, 120) + para('MARK T1 line 360', ppr=spacing(after=0, line=360, rule='auto'), run_rpr=T))
style = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr>' + spacing(after=0, line=240, rule='auto') + '</w:pPr></w:style>'
write(out_path(), body, styles_xml(doc_defaults('Times New Roman', 24)).replace('<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>', style), settings_xml())
