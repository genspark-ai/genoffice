import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
T = rpr('Times New Roman', 24)
normal = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:ind w:right="142" w:firstLine="700"/></w:pPr></w:style>'
hang = '<w:style w:type="paragraph" w:styleId="Hang"><w:name w:val="Hang"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="0" w:hanging="400"/></w:pPr></w:style>'
styles = styles_xml(doc_defaults('Times New Roman', 24), hang).replace('<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>', normal)
levels = [[lvl(ind='<w:ind w:left="0"/>')], [lvl(ind='<w:ind w:left="720"/>')], [lvl(ind='<w:ind w:left="720" w:hanging="360"/>')],
          [lvl(ind='<w:ind w:left="720" w:firstLine="0"/>')], [lvl(ind='<w:ind w:left="0" w:hanging="360"/>')], [lvl(ind='<w:ind w:left="0"/>', rpr='<w:rPr><w:sz w:val="36"/></w:rPr>')]]
def item(num, text, ppr=''):
    return f'<w:p><w:pPr>{ppr}<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num}"/></w:numPr></w:pPr><w:r>{T}<w:t>{text}</w:t></w:r></w:p>'
body = (item(1, 'A lvl left=0 only') + item(2, 'B lvl left=720 only') + item(3, 'C lvl left=720 hanging=360') + item(4, 'D lvl left=720 firstLine=0')
        + item(2, 'E style hanging=400 + lvl left=720', '<w:pStyle w:val="Hang"/>') + item(5, 'F para left=1440 + lvl left=0 hanging=360', '<w:ind w:left="1440"/>')
        + item(1, 'G para firstLine=0', '<w:ind w:firstLine="0"/>') + item(1, 'H para hanging=200', '<w:ind w:hanging="200"/>') + item(6, 'K oversized 18pt number'))
write(out_path(), body, styles, settings_xml(), numbering=numbering_xml(levels))
