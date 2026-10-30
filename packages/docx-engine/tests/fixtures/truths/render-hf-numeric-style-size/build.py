import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Letter margins 720, header 720, docDefaults Calibri 12 single. Two nextPage sections, one
# header each: (1) a paragraph in style "52" (name header, rPr Malgun Gothic sz 20) whose
# run names no size or face; (2) control: the same run with its own sz 20 and Malgun face.
C = rpr('Calibri', 24)
SP = spacing(after=0, line=240, rule='auto')
PAGE = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>'
KO = '페이지 1'
MALGUN = 'w:ascii="Malgun Gothic" w:hAnsi="Malgun Gothic" w:eastAsia="Malgun Gothic" w:cs="Malgun Gothic"'
STYLES = (
    '<w:style w:type="paragraph" w:styleId="52"><w:name w:val="header"/><w:basedOn w:val="Normal"/>'
    f'<w:rPr><w:rFonts {MALGUN}/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>'
)
HEADERS = {
    'header1.xml': hdr_xml(f'<w:p><w:pPr><w:pStyle w:val="52"/></w:pPr><w:r><w:t>{KO}</w:t></w:r></w:p>'),
    'header2.xml': hdr_xml(f'<w:p><w:r><w:rPr><w:rFonts {MALGUN}/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t>{KO}</w:t></w:r></w:p>'),
}


def sect(n):
    return f'<w:sectPr><w:headerReference w:type="default" r:id="rIdheader{n}"/>{PAGE}</w:sectPr>'


body = para('Section 1 body', ppr=SP + sect(1), run_rpr=C, mark_rpr='') + para('Section 2 body', ppr=SP, run_rpr=C)
write(out_path(), body, styles_xml(doc_defaults('Calibri', 24), STYLES), settings_xml(), sect_pr=sect(2), hf=HEADERS)
