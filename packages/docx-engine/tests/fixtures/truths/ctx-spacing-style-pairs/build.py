import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
SP0 = spacing(before=0, after=0)
H1 = ('<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>'
      '<w:pPr>' + SP0 + '<w:outlineLvl w:val="0"/></w:pPr></w:style>')
LP = ('<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>'
      '<w:pPr><w:ind w:left="720"/>' + spacing(after=0) + '</w:pPr></w:style>')
NUM = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
CTX = '<w:contextualSpacing/>'
def p(text, ppr=''):
    return para(text, ppr=ppr)
body = ''.join([
    p('control one', SP0),
    p('control two', SP0),
    p('outline heading', SP0 + '<w:outlineLvl w:val="0"/>'),
    p('item after outline heading', NUM + spacing(before=240, after=0) + CTX),
    p('body', SP0),
    p('styled heading', '<w:pStyle w:val="Heading1"/>'),
    p('item after styled heading', NUM + spacing(before=240, after=0) + CTX),
    p('body', SP0),
    p('plain normal', SP0),
    p('item after plain normal', NUM + spacing(before=240, after=0) + CTX),
    p('body', SP0),
    p('lp item one', '<w:pStyle w:val="ListParagraph"/>' + NUM + spacing(before=240, after=160) + CTX),
    p('lp item two', '<w:pStyle w:val="ListParagraph"/>' + NUM + spacing(before=240, after=160) + CTX),
    p('body after lp', SP0),
    p('ctx para before outline heading', spacing(before=0, after=240) + CTX),
    p('outline heading after ctx para', SP0 + '<w:outlineLvl w:val="0"/>'),
    p('body', SP0),
    p('ctx para before styled heading', spacing(before=0, after=240) + CTX),
    p('styled heading after ctx para', '<w:pStyle w:val="Heading1"/>'),
    p('end', SP0),
])
num = numbering_xml([[lvl('bullet', '-', ind='<w:ind w:left="720" w:hanging="360"/>')]])
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22), H1 + LP), settings_xml(), numbering=num)
