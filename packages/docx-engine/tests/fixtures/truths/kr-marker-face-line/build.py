import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Korean Word template (theme minor Latin = Malgun Gothic, lang eastAsia ko-KR):
# body style in Gulim, list style Times New Roman + eastAsia Batang, bullet
# levels whose rFonts name Malgun Gothic / Batang (hint eastAsia) or Times New Roman
MAL = '맑은 고딕'
GUL = '굴림'
BAT = '바탕'
TNR = 'Times New Roman'
dd = ('<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>'
      '<w:kern w:val="2"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ko-KR" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
      f'<w:pPrDefault><w:pPr>{spacing(after=160, line=259, rule="auto")}<w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>')
st = (
  '<w:style w:type="paragraph" w:default="1" w:styleId="a"><w:name w:val="Normal"/><w:pPr><w:widowControl w:val="0"/><w:wordWrap w:val="0"/><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/></w:pPr></w:style>'
  '<w:style w:type="paragraph" w:styleId="a6"><w:name w:val="List Paragraph"/><w:basedOn w:val="a"/><w:pPr><w:widowControl/><w:wordWrap/><w:autoSpaceDE/><w:autoSpaceDN/>'
  f'{spacing(after=0, line=240, rule="auto")}<w:ind w:leftChars="400" w:left="800"/><w:jc w:val="left"/></w:pPr>'
  f'<w:rPr><w:rFonts w:ascii="{TNR}" w:eastAsia="{BAT}" w:hAnsi="{TNR}" w:cs="{TNR}"/><w:kern w:val="0"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:eastAsia="it-IT"/></w:rPr></w:style>'
  '<w:style w:type="paragraph" w:styleId="a7"><w:name w:val="Normal (Web)"/><w:basedOn w:val="a"/><w:pPr><w:widowControl/><w:wordWrap/><w:autoSpaceDE/><w:autoSpaceDN/>'
  '<w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1" w:line="240" w:lineRule="auto"/><w:jc w:val="left"/></w:pPr>'
  f'<w:rPr><w:rFonts w:ascii="{GUL}" w:eastAsia="{GUL}" w:hAnsi="{GUL}" w:cs="{GUL}"/><w:kern w:val="0"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>'
)
styles = f'{XML}<w:styles {W}>{dd}{st}</w:styles>'
def mk(face, hint):
    h = f' w:hint="{hint}"' if hint else ''
    return f'<w:rPr><w:rFonts w:ascii="{face}" w:eastAsia="{face}" w:hAnsi="{face}" w:cs="{TNR}"{h}/></w:rPr>'
ind = '<w:ind w:left="760" w:hanging="360"/>'
levels = [
  [lvl('bullet', '&#x2022;', ind, mk(MAL, 'eastAsia'))],
  [lvl('bullet', '&#x2022;', ind, mk(BAT, 'eastAsia'))],
  [lvl('bullet', '&#x2022;', ind, mk(TNR, 'default'))],
  [lvl('bullet', '&#x2022;', ind, mk(GUL, 'eastAsia'))],
]
LAT = 'Tomorrow planning review 0123 journal'
tnr22 = f'<w:rPr><w:rFonts w:ascii="{TNR}" w:hAnsi="{TNR}" w:cs="{TNR}"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'
sz22 = '<w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'
cal = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr>'
a7 = '<w:pStyle w:val="a7"/><w:spacing w:before="120" w:beforeAutospacing="0" w:after="120" w:afterAutospacing="0"/><w:jc w:val="both"/>'
def lines(tag, n=4): return [f'{tag}{i} {LAT}' for i in range(n)]
single = spacing(after=0, line=240, rule='auto')
body = para(lines('C'), ppr=single, run_rpr=cal)      # 0 control Calibri 13.44
body += para(lines('P1'), ppr=a7, run_rpr=tnr22)       # 1 Gulim style, Times New Roman 11 runs: 12.72
body += para(lines('P3'), ppr=a7, run_rpr=sz22)        # 2 Gulim 11 runs: 14.16-14.40
for n in range(1, 5):                                  # 3.. items (blocks 3-6, 8-11, 13-16, 18-21)
    ppr = f'<w:pStyle w:val="a6"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{n}"/></w:numPr><w:ind w:leftChars="0" w:hanging="357"/>'
    for i in range(3):
        body += para(f'B{n}.{i} {LAT}', ppr=ppr, run_rpr=sz22)
    body += para(lines(f'B{n}w', 3), ppr=ppr, run_rpr=sz22)
    body += para(f'X{n} {LAT}', ppr=single, run_rpr=cal)
ft = font_table_xml([
  (MAL, {'panose1': '020B0503020000020004', 'charset': '81', 'family': 'modern', 'pitch': 'variable'}),
  (GUL, {'altName': 'Gulim', 'panose1': '020B0600000101010101', 'charset': '81', 'family': 'modern', 'pitch': 'variable'}),
  (BAT, {'altName': 'Batang', 'panose1': '02030600000101010101', 'charset': '81', 'family': 'roman', 'pitch': 'variable'}),
  (TNR, {'panose1': '02020603050405020304', 'charset': '00', 'family': 'roman', 'pitch': 'variable'}),
  ('Calibri', {'panose1': '020F0502020204030204', 'charset': '00', 'family': 'swiss', 'pitch': 'variable'}),
])
write(out_path(), body, styles, settings_xml(), font_table=ft, sect=A4, numbering=numbering_xml(levels), theme=theme_xml(minor=MAL, minor_ea=''))
