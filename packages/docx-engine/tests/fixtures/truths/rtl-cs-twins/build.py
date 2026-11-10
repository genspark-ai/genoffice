import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
h1 = '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
bt = '<w:style w:type="paragraph" w:styleId="BodyText"><w:name w:val="Body Text"/><w:basedOn w:val="Normal"/><w:rPr><w:szCs w:val="24"/></w:rPr></w:style>'
dd = ('<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
      '<w:pPrDefault><w:pPr>' + spacing(after=0, line=240, rule='auto') + '</w:pPr></w:pPrDefault></w:docDefaults>')
AR = 'عنوان التقرير'
body = ('<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:bidi/></w:pPr><w:r><w:rPr><w:rtl/></w:rPr><w:t>' + AR + '</w:t></w:r></w:p>'
        + '<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:bidi/>' + spacing(after=0, line=360, rule='auto') + '<w:rPr><w:sz w:val="36"/><w:rtl/></w:rPr></w:pPr></w:p>')
write(out_path(), body, styles_xml(dd, h1 + bt), settings_xml())
