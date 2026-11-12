import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *


def tbl(tcws, rows):
    g = ''.join(f'<w:gridCol w:w="{w}"/>' for w in tcws)
    trs = ''.join(
        '<w:tr>' + ''.join(
            f'<w:tc><w:tcPr><w:tcW w:w="{tcws[i]}" w:type="dxa"/></w:tcPr>{para(t)}</w:tc>'
            for i, t in enumerate(r)) + '</w:tr>'
        for r in rows)
    return f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>{g}</w:tblGrid>{trs}</w:tbl>'


WRAP = [
    ['#', 'ID', 'Descricao', 'Entrada', 'Saida', 'Dono'],
    ['1', 'R1', 'Esta etapa consiste em analisar os requisitos do sistema e registar cada caso', 'Documento de requisitos aprovado pelo cliente', 'Lista', 'Analista'],
    ['2', 'R2', 'Registo das regras de negocio junto ao cliente e registo em ata', 'Ata', 'Lista', 'Gerente'],
]
narrow = lambda text: ([600, 8400], [[text, 'filler text that wraps across a few lines so the wide column has surplus']])
body = (para('A') + tbl([1504] * 6, WRAP)
        + para('B') + tbl([1000] * 6, [['#', 'ID', 'Nome', 'Tipo', 'Estado', 'Dono'], ['1', 'R1', 'Login', 'Documento', 'Ok', 'Ana']])
        + para('C') + tbl(*narrow('mmmm · mmmm'))
        + para('D') + tbl(*narrow('mmmm mmmm'))
        + para('E') + tbl(*narrow('mmmm/mmmm'))
        + para('F') + tbl(*narrow('mmmm-mmmm'))
        + para('G') + tbl(*narrow('mmmm : mmmm'))
        + para('H'))
A4 = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
write(out_path(), body, styles_xml(doc_defaults()), settings_xml(), sect=A4)
