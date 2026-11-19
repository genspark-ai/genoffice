import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Letter, docDefaults Calibri 11 + eastAsia Microsoft YaHei, line=300 auto, compat 14, no typed grid. CJK runs
# declaring two faces (ascii Calibri + eastAsia Microsoft YaHei) at 8.5pt and 10.5pt,
# a single-face control (every slot Microsoft YaHei), then the same lines inside
# 1-column bordered tables with zero top/bottom cell margins
HAN = '文字段落行高测试'
SP = spacing(after=0, line=300, rule='auto')
C = rpr('Calibri', 22)


def two(sz):
    return rpr('Calibri', sz, east_asia='Microsoft YaHei')


def one(sz):
    return rpr('Microsoft YaHei', sz, east_asia='Microsoft YaHei')


def table(rp, n=3):
    b = ''.join(
        f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')
    )
    mar = '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
    row = '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' + para(HAN, ppr=SP, run_rpr=rp) + '</w:tc></w:tr>'
    return (
        f'<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders>{b}</w:tblBorders>{mar}</w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' + row * n + '</w:tbl>'
    )


body = (
    ''.join(para(HAN, ppr=SP, run_rpr=two(17)) for _ in range(3))
    + ''.join(para(HAN, ppr=SP, run_rpr=two(21)) for _ in range(3))
    + ''.join(para(HAN, ppr=SP, run_rpr=one(17)) for _ in range(3))
    + para('MARK', ppr=SP, run_rpr=C)
    + table(two(17))
    + para('MID', ppr=SP, run_rpr=C)
    + table(two(21))
    + para('MID2', ppr=SP, run_rpr=C)
    + table(one(17))
    + para('END', ppr=SP, run_rpr=C)
)
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22, east_asia='Microsoft YaHei', ppr=SP)), settings_xml(compat_mode=14))
