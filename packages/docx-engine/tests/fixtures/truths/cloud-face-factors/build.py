import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
# body-only twin of render-cloud-face-pitch: one paragraph per face/spacing
# (Dubai 12pt, Algerian 18pt, Google Sans Text 11pt via fontTable altName Calibri)
AR = 'النص العربي لقياس السطر'
SP = spacing(after=0, line=240, rule='auto')
SP276 = spacing(after=0, line=276, rule='auto')
SIG = '<w:sig w:usb0="E00002FF" w:usb1="4000A04A" w:usb2="00000008" w:usb3="00000000" w:csb0="00000041" w:csb1="00000000"/>'
FONTS = (
    f'{XML}<w:fonts {W}>'
    '<w:font w:name="Dubai"><w:panose1 w:val="020B0503030200000000"/><w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/>' + SIG + '</w:font>'
    '<w:font w:name="Google Sans Text"><w:altName w:val="Calibri"/><w:charset w:val="00"/><w:family w:val="auto"/><w:pitch w:val="variable"/></w:font>'
    '<w:font w:name="Algerian"><w:panose1 w:val="04020705040A02060702"/><w:charset w:val="00"/><w:family w:val="decorative"/><w:pitch w:val="variable"/></w:font>'
    '</w:fonts>'
)
body = (
    para(AR, ppr=SP, run_rpr=rpr('Dubai', 24))
    + para('ALGERIAN SAMPLE LINE', ppr=SP, run_rpr=rpr('Algerian', 36))
    + para('Google Sans Text sample line', ppr=SP, run_rpr=rpr('Google Sans Text', 22))
    + para('Google Sans Text at 276', ppr=SP276, run_rpr=rpr('Google Sans Text', 22))
    + para(AR, ppr=SP276, run_rpr=rpr('Dubai', 24))
)
write(out_path(), body, styles_xml(doc_defaults('Calibri', 22)), settings_xml(), font_table=FONTS)
