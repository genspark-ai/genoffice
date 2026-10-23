import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
A = rpr('Arial', 22)
SP = spacing(after=0, line=240, rule='auto')
H1 = '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:keepLines/>' + spacing(before=240, after=80) + '</w:pPr><w:rPr><w:sz w:val="32"/></w:rPr></w:style>'
filler = para([f'filler {i+1}' for i in range(31)], ppr=SP, run_rpr=A)
heads = ''.join('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:p>' for _ in range(11))
body = filler + heads + para(None, ppr=SP, run_rpr=A) + para('anchor text', ppr=SP, run_rpr=A)
sect = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>'
write(out_path(), body, styles_xml(doc_defaults('Arial', 22), H1), settings_xml(), sect=sect)
