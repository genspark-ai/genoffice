import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Letter margins 720, header 720, docDefaults Calibri 12 single. Two nextPage sections, one
# header each, both paragraphs in a style with sz 32 szCs 22 and cs Times New Roman:
# (1) the only run is neutral characters (em dashes, a digit) under w:rtl; (2) control:
# the same run without w:rtl.
C = rpr('Calibri', 24)
SP = spacing(after=0, line=240, rule='auto')
PAGE = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>'
STYLES = (
    '<w:style w:type="paragraph" w:styleId="HeaderRtl"><w:name w:val="header rtl"/><w:basedOn w:val="Normal"/>'
    '<w:rPr><w:rFonts w:cs="Times New Roman"/><w:sz w:val="32"/><w:szCs w:val="22"/></w:rPr></w:style>'
)
TEXT = '<w:t xml:space="preserve">— 1 —</w:t>'
HEADERS = {
    'header1.xml': hdr_xml(f'<w:p><w:pPr><w:pStyle w:val="HeaderRtl"/></w:pPr><w:r><w:rPr><w:rtl/></w:rPr>{TEXT}</w:r></w:p>'),
    'header2.xml': hdr_xml(f'<w:p><w:pPr><w:pStyle w:val="HeaderRtl"/></w:pPr><w:r>{TEXT}</w:r></w:p>'),
}


def sect(n):
    return f'<w:sectPr><w:headerReference w:type="default" r:id="rIdheader{n}"/>{PAGE}</w:sectPr>'


body = para('Section 1 body', ppr=SP + sect(1), run_rpr=C, mark_rpr='') + para('Section 2 body', ppr=SP, run_rpr=C)
write(out_path(), body, styles_xml(doc_defaults('Calibri', 24), STYLES), settings_xml(), sect_pr=sect(2), hf=HEADERS)
