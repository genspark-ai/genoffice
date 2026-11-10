import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
styles = styles_xml('<w:docDefaults><w:rPrDefault/><w:pPrDefault/></w:docDefaults>')
P = '<w:pBdr><w:top w:val="none" w:sz="0" w:space="30" w:color="auto"/><w:left w:val="none" w:sz="0" w:space="31" w:color="auto"/><w:bottom w:val="none" w:sz="0" w:space="30" w:color="auto"/><w:right w:val="none" w:sz="0" w:space="31" w:color="auto"/></w:pBdr>'
body = (para(['Hamburgefonstiv 123 quick', 'Hamburgefonstiv 123 quick', 'Hamburgefonstiv 123 quick'])
        + para('padded paragraph', ppr=P)
        + para('方正 probe', run_rpr=rpr('方正小标宋简体', 24, east_asia='方正小标宋简体')))
ft = font_table_xml([('方正小标宋简体', {'altName': '宋体', 'charset': '86', 'family': 'auto'})])
write(out_path(), body, styles, settings_xml(), font_table=ft)
