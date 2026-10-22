import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
def st(sid, rfonts):
    rf = f'<w:rFonts {rfonts}/>' if rfonts else ''
    return f'<w:style w:type="paragraph" w:styleId="{sid}"><w:name w:val="{sid}"/><w:basedOn w:val="Normal"/><w:rPr>{rf}<w:sz w:val="28"/></w:rPr></w:style>'
styles = styles_xml(doc_defaults('Arial', 22), st('HA', 'w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"') + st('HB', 'w:ascii="Georgia" w:hAnsi="Georgia"') + st('HC', '') + st('HD', 'w:ascii="Georgia" w:hAnsi="Georgia" w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"'))
def h(sid, text):
    return f'<w:p><w:pPr><w:pStyle w:val="{sid}"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'
body = h('HA', 'HA theme major style') + h('HB', 'HB explicit Georgia style') + h('HC', 'HC no rFonts style') + h('HD', 'HD Georgia plus theme major') + para('Normal body control')
write(out_path(), body, styles, settings_xml(), sect=A4, theme=theme_xml(major='Calibri', minor='Cambria'))
