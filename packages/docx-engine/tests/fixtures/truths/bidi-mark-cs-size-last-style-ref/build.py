import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
SZCS = '<w:szCs w:val="22"/>'
defaults = (
    f'<w:docDefaults><w:rPrDefault>{rpr("Times New Roman", 22, extra=SZCS)}</w:rPrDefault>'
    f'<w:pPrDefault><w:pPr>{spacing(after=0, line=240, rule="auto")}</w:pPr></w:pPrDefault></w:docDefaults>'
)
styles = (
    f'{XML}<w:styles {W}>{defaults}'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="24"/></w:rPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:sz w:val="32"/></w:rPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="MdHeading1"><w:name w:val="MdHeading1"/><w:pPr><w:spacing w:before="480"/></w:pPr><w:rPr><w:b/><w:sz w:val="36"/></w:rPr></w:style>'
    '<w:style w:type="character" w:styleId="Red"><w:name w:val="Red"/><w:rPr><w:color w:val="FF0000"/></w:rPr></w:style>'
    '<w:style w:type="character" w:styleId="Blue"><w:name w:val="Blue"/><w:rPr><w:color w:val="0000FF"/></w:rPr></w:style>'
    '</w:styles>'
)
def empty(mark, bidi=True):
    return para(None, ppr='<w:bidi/>' if bidi else '', mark_rpr=f'<w:rPr>{mark}</w:rPr>')
body = (
    empty('<w:noProof/>')
    + empty('<w:noProof/>', bidi=False)
    + empty('<w:sz w:val="36"/>')
    + empty('<w:rtl w:val="0"/>')
    + empty('<w:szCs w:val="40"/>')
    + '<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:pStyle w:val="MdHeading1"/></w:pPr>'
    '<w:r><w:rPr><w:rStyle w:val="Red"/><w:rStyle w:val="Blue"/></w:rPr><w:t>Title</w:t></w:r></w:p>'
)
write(out_path(), body, styles, settings_xml())
