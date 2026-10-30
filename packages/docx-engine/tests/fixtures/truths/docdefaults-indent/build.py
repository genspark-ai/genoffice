import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# docDefaults pPrDefault w:ind (left 567, right 567, firstLine 720) under a
# Normal with no pPr: body paragraphs, a Normal-based style without w:ind, a root
# style without w:basedOn, a style cancelling firstLine with w:firstLine="0", a
# paragraph with a direct w:ind left="0", and a hanging=360 style.
ppr = spacing(after=0, line=240, rule='auto') + '<w:ind w:left="567" w:right="567" w:firstLine="720"/>'
styles = styles_xml(
    doc_defaults('Calibri', 22, ppr=ppr),
    '<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/><w:basedOn w:val="Normal"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="Root"><w:name w:val="Root"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="NoFirst"><w:name w:val="No First"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:firstLine="0"/></w:pPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="Hang"><w:name w:val="Hang"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:hanging="360"/></w:pPr></w:style>',
)
long = 'alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau upsilon phi chi psi omega ' * 3
body = para(long)
body += para(long, ppr='<w:pStyle w:val="Body"/>')
body += para(long, ppr='<w:pStyle w:val="Root"/>')
body += para(long, ppr='<w:pStyle w:val="NoFirst"/>')
body += para(long, ppr='<w:ind w:left="0"/>')
body += para(long, ppr='<w:pStyle w:val="Hang"/>')
body += para(long, ppr='<w:ind w:firstLine="0"/>')
write(out_path(), body, styles, settings_xml())
