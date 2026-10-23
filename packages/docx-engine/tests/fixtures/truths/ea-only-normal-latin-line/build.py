import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
normal = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:jc w:val="both"/></w:pPr><w:rPr><w:rFonts w:eastAsiaTheme="minorHAnsi" w:cstheme="minorBidi"/><w:sz w:val="24"/><w:lang w:val="zh-CN"/></w:rPr></w:style>'
styles = styles_xml(doc_defaults('Times New Roman', 24, east_asia='SimSun')).replace('<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>', normal)
body = para(['Latin only line one of the probe', 'Latin only line two', 'Latin only line three'], ppr=spacing(after=0, line=240, rule='auto'))
write(out_path(), body, styles, settings_xml('<w:useFELayout/>'), sect=A4, theme=theme_xml(major='Aptos Display', minor='Aptos'))
