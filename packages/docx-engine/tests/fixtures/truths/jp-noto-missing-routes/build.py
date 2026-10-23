import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
N = 'Noto Sans CJK JP'
normal = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="' + N + '" w:hAnsi="' + N + '" w:eastAsia="' + N + '" w:cs="' + N + '"/><w:sz w:val="21"/></w:rPr></w:style>'
styles = styles_xml(doc_defaults('Calibri', 21, ppr=spacing(after=0, line=240, rule='auto'))).replace('<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>', normal)
body = para('あいうえお kana line') + para('Latin digits 2026')
write(out_path(), body, styles, settings_xml('<w:useFELayout/>', 14), sect=A4)
