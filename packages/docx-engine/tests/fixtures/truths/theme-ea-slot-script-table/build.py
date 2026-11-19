import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# Word probe 2026-11-03 p1 (ko) / p2 (ja) / p4 (ko, no script table): the script table's
# face for themeFontLang eastAsia beats the literal <a:ea> typeface
HG = '한글 본문 테스트 문장입니다'
KANA = 'あいうえお 日本語の本文'
BATANG = '바탕'; MALGUN = '맑은 고딕'; MSMIN = 'ＭＳ 明朝'; YUMIN = '游明朝'
MINOR3 = '<w:rFonts w:asciiTheme="minorEastAsia" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorEastAsia"/>'
MINOREA = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsiaTheme="minorEastAsia"/>'
MAJOR3 = '<w:rFonts w:asciiTheme="majorEastAsia" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorEastAsia"/>'
lit = lambda f: f'<w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:eastAsia="{f}" w:cs="{f}"/>'
MARK = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/><w:sz w:val="24"/><w:lang w:eastAsia="en-US"/>'
def theme(ea, scripts):
    A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    sc = ''.join(f'<a:font script="{k}" typeface="{v}"/>' for k, v in scripts.items())
    grp = lambda tag: f'<a:{tag}><a:latin typeface="Calibri"/><a:ea typeface="{ea}"/><a:cs typeface=""/>{sc}</a:{tag}>'
    return f'{XML}<a:theme {A} name="Office"><a:themeElements><a:clrScheme name="Office"/><a:fontScheme name="Office">{grp("majorFont")}{grp("minorFont")}</a:fontScheme><a:fmtScheme name="Office"/></a:themeElements></a:theme>'
def settings(lang):
    return settings_xml().replace('<w:compat>', f'<w:themeFontLang w:val="en-US" w:eastAsia="{lang}"/><w:compat>')
def marker(m): return para([f'#M:{m}#'], ppr=spacing(before=0, after=0, line=240, rule='auto'), run_rpr=f'<w:rPr>{MARK}</w:rPr>', mark_rpr=f'<w:rPr>{MARK}</w:rPr>')
def block(rf, text, lines=4): return para([text] * lines, ppr=spacing(before=0, after=0, line=240, rule='auto'), run_rpr=f'<w:rPr>{rf}<w:sz w:val="24"/></w:rPr>', mark_rpr=f'<w:rPr>{rf}<w:sz w:val="24"/></w:rPr>')
dd = '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:sz w:val="24"/><w:lang w:val="en-US" w:eastAsia="%s" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>%s</w:pPr></w:pPrDefault></w:docDefaults>'
out = out_path(); d = os.path.dirname(out)
ko = marker('A') + block(MINOR3, HG) + marker('B') + block(MINOREA, HG) + marker('C') + block(MINOR3, '0123456789 abc') + marker('D') + block(MAJOR3, HG) + marker('E') + block(lit(BATANG), HG) + marker('F') + block(lit(MALGUN), HG) + marker('G')
write(os.path.join(d, 'probe.docx'), ko, styles_xml(dd % ('ko-KR', spacing(after=0, line=240, rule='auto'))), settings('ko-KR'), theme=theme(BATANG, {'Jpan': YUMIN, 'Hang': MALGUN}))
write(os.path.join(d, 'probe-reverse.docx'), marker('A') + block(MINOR3, HG) + marker('B'), styles_xml(dd % ('ko-KR', spacing(after=0, line=240, rule='auto'))), settings('ko-KR'), theme=theme(MALGUN, {'Jpan': YUMIN, 'Hang': BATANG}))
write(os.path.join(d, 'probe-notable.docx'), marker('A') + block(MINOR3, HG) + marker('B'), styles_xml(dd % ('ko-KR', spacing(after=0, line=240, rule='auto'))), settings('ko-KR'), theme=theme(BATANG, {}))
ja = marker('A') + block(MINOR3, KANA) + marker('B') + block(MINOREA, KANA) + marker('C') + block(MINOR3, HG) + marker('D') + block(lit(MSMIN), KANA) + marker('E') + block(lit(YUMIN), KANA) + marker('F')
write(os.path.join(d, 'probe-ja.docx'), ja, styles_xml(dd % ('ja-JP', spacing(after=0, line=240, rule='auto'))), settings('ja-JP'), theme=theme(MSMIN, {'Jpan': YUMIN, 'Hang': MALGUN}))
