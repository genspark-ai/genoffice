import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Letter, docDefaults Times New Roman sz 19 (9.5pt), line 276 auto. A bordered two-column
# table, cell paragraphs before/after 40; the left cell of each row holds one oMathPara in
# Cambria Math: a sub-only line, a sum with undOvr limits, a display fraction. The right
# cell is a one-line label. A body paragraph closes the document.
T = rpr('Times New Roman', 19)
CM = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>'
SP = spacing(before=40, after=40, line=276, rule='auto')
BORD = ''.join(
    f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')
)


def mr(text):
    return f'<m:r>{CM}<m:t xml:space="preserve">{text}</m:t></m:r>'


def sub(base, script):
    return f'<m:sSub><m:e>{mr(base)}</m:e><m:sub>{mr(script)}</m:sub></m:sSub>'


EQ_SUB = sub('P', 't|t') + mr('=(I−') + sub('K', 't') + mr('H)') + sub('P', 't|t−1')
EQ_NARY = (
    '<m:nary><m:naryPr><m:chr m:val="∑"/><m:limLoc m:val="undOvr"/></m:naryPr>'
    f'<m:sub>{mr("i=1")}</m:sub><m:sup>{mr("n")}</m:sup><m:e>{sub("x", "i")}</m:e></m:nary>'
    + mr('=μ')
)
EQ_FRAC = f'<m:f><m:num>{mr("a+b")}</m:num><m:den>{mr("c")}</m:den></m:f>' + mr('=d')


def math_para(inner):
    return f'<w:p><w:pPr>{SP}</w:pPr><m:oMathPara {M}><m:oMath>{inner}</m:oMath></m:oMathPara></w:p>'


def row(eq, label):
    return (
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>' + math_para(eq) + '</w:tc>'
        '<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>' + para(label, ppr=SP, run_rpr=T) + '</w:tc></w:tr>'
    )


table = (
    f'<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders>{BORD}</w:tblBorders></w:tblPr>'
    '<w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>'
    + row(EQ_SUB, 'sub-only line')
    + row(EQ_NARY, 'nary with undOvr limits')
    + row(EQ_FRAC, 'display fraction')
    + '</w:tbl>'
)
body = table + para('tail', ppr=spacing(after=0, line=276, rule='auto'), run_rpr=T)
write(out_path(), body, styles_xml(doc_defaults('Times New Roman', 19, ppr=spacing(after=0, line=276, rule='auto'))), settings_xml())
