import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Word probe 2026-11-03 p5 / p6: docDefaults w:lang w:val=ko-KR sizes empty and
# space-only paragraphs by the eastAsia face (Malgun 17.28 at 10pt), w:val=en-US by Calibri
HG = '한글 본문 테스트 문장입니다'
MARK = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/><w:sz w:val="24"/><w:lang w:eastAsia="en-US"/>'
SP = spacing(before=0, after=0, line=240, rule='auto')
def marker(m): return para([f'#M:{m}#'], ppr=SP, run_rpr=f'<w:rPr>{MARK}</w:rPr>', mark_rpr=f'<w:rPr>{MARK}</w:rPr>')
EMPTY = f'<w:p><w:pPr>{SP}</w:pPr></w:p>'
EMPTY600 = f'<w:p><w:pPr>{spacing(before=0, after=0, line=600, rule="auto")}</w:pPr></w:p>'
SPACE = f'<w:p><w:pPr>{SP}</w:pPr><w:r><w:t xml:space="preserve">   </w:t></w:r></w:p>'
TEXTKO = f'<w:p><w:pPr>{SP}</w:pPr><w:r><w:t>{HG}</w:t></w:r></w:p>'
TEXTLA = f'<w:p><w:pPr>{SP}</w:pPr><w:r><w:t>Latin only text line</w:t></w:r></w:p>'
body = marker('A') + EMPTY * 3 + marker('B') + SPACE * 3 + marker('C') + TEXTKO * 3 + marker('D') + TEXTLA * 3 + marker('E') + EMPTY600 * 2 + marker('F')
dd = '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Malgun Gothic" w:hAnsi="Calibri"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="%s" w:eastAsia="%s"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>%s</w:pPr></w:pPrDefault></w:docDefaults>'
out = out_path(); d = os.path.dirname(out)
write(os.path.join(d, 'probe.docx'), body, styles_xml(dd % ('ko-KR', 'ko-KR', SP)), settings_xml())
write(os.path.join(d, 'probe-en.docx'), body, styles_xml(dd % ('en-US', 'en-US', SP)), settings_xml())
