import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
T = rpr(sz=24)
AUTO = ' w:beforeAutospacing="1" w:afterAutospacing="1"'
NW = ('<w:style w:type="paragraph" w:styleId="NW"><w:name w:val="NW"/><w:basedOn w:val="Normal"/>'
      f'<w:pPr>{spacing(before=100, after=100, extra=AUTO)}</w:pPr></w:style>'
      '<w:style w:type="paragraph" w:styleId="NW2"><w:name w:val="NW2"/><w:basedOn w:val="NW"/>'
      f'<w:pPr>{spacing(before=0, after=0)}</w:pPr></w:style>')
def n(t): return para(t, run_rpr=T)
def s(t, style, sp=''): return para(t, ppr=f'<w:pStyle w:val="{style}"/>{sp}', run_rpr=T)
body = ''.join([
    n('N0'),
    s('T1 style auto', 'NW'), n('N1'),
    s('T2 literal 0 no attr', 'NW', spacing(before=0, after=0)), n('N2'),
    s('T3 literal 0 attr 0', 'NW', spacing(before=0, after=0, extra=' w:beforeAutospacing="0" w:afterAutospacing="0"')), n('N3'),
    s('T4 literal 400 no attr', 'NW', spacing(before=400, after=400)), n('N4'),
    s('T5 direct auto', 'Normal', spacing(before=0, after=0, extra=AUTO)), n('N5'),
    s('T6 NW2 style literal 0', 'NW2'), n('N6'),
    s('T7 literal 0 line 480 no attr', 'NW', spacing(before=0, after=0, line=480, rule='auto')), n('N7'),
])
mode = int(sys.argv[2]) if len(sys.argv) > 2 else 15
write(out_path(), body, styles_xml(doc_defaults('Times New Roman', 24), NW), settings_xml('', mode))
