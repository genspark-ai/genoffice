import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Letter margins 720, header 720, Calibri 11 single. Section 1 (portrait) declares a
# header holding a 5000-pct two-column table whose right cell wraps its title to two
# lines at the portrait width; section 2 (landscape) has no headerReference and must
# draw that header at the portrait width (two lines, same body top); section 3
# (landscape) declares its own copy of the header, which spans the landscape column
# and wraps to one line.
C = rpr('Calibri', 22)
SP = spacing(after=0, line=240, rule='auto')
PORTRAIT = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>'
LANDSCAPE = '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>'
TITLE = 'DOCUMENT EVALUATION DU PROGRAMME DE FORMATION'


def header():
    cell = lambda text: (
        '<w:tc><w:tcPr><w:tcW w:w="2500" w:type="pct"/></w:tcPr>' + para(text, ppr=SP, run_rpr=C) + '</w:tc>'
    )
    return hdr_xml(
        '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="5400"/><w:gridCol w:w="5400"/></w:tblGrid>'
        '<w:tr>' + cell('LOGO') + cell(TITLE) + '</w:tr></w:tbl>' + para(None, ppr=SP, run_rpr=C)
    )


def sect(page, header_ref=None):
    ref = f'<w:headerReference w:type="default" r:id="rId{header_ref}"/>' if header_ref else ''
    return f'<w:sectPr>{ref}{page}</w:sectPr>'


body = (
    para('Section one body', ppr=SP + sect(PORTRAIT, 'header1'), run_rpr=C, mark_rpr='')
    + para('Section two body', ppr=SP + sect(LANDSCAPE), run_rpr=C, mark_rpr='')
    + para('Section three body', ppr=SP, run_rpr=C)
)
write(
    out_path(),
    body,
    styles_xml(doc_defaults('Calibri', 22)),
    settings_xml(),
    sect_pr=sect(LANDSCAPE, 'header2'),
    hf={'header1.xml': header(), 'header2.xml': header()},
)
