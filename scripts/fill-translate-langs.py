#!/usr/bin/env python3
"""Write the missing translate-target language names into every locale file.

The menus are driven by a key list, not by the locale files, so a missing
entry renders the raw key name instead of failing the build. This fills them
in from one table so the three i18n trees cannot drift again.
"""
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

# suffix -> the 14 languages that were missing (Traditional Chinese is in
# here too: the docs trees never had it, and slides already did)
NEW = ['Thai', 'Indonesian', 'Russian', 'Arabic', 'Portuguese', 'Italian',
       'Polish', 'Czech', 'Dutch', 'Malay', 'Hebrew', 'Hindi',
       'TraditionalChinese', 'Vietnamese']

T = {
 'ar': {'Thai':'التايلاندية','Indonesian':'الإندونيسية','Russian':'الروسية','Arabic':'العربية','Portuguese':'البرتغالية','Italian':'الإيطالية','Polish':'البولندية','Czech':'التشيكية','Dutch':'الهولندية','Malay':'الملايوية','Hebrew':'العبرية','Hindi':'الهندية','TraditionalChinese':'الصينية التقليدية','Vietnamese':'الفيتنامية'},
 'cs': {'Thai':'thajština','Indonesian':'indonéština','Russian':'ruština','Arabic':'arabština','Portuguese':'portugalština','Italian':'italština','Polish':'polština','Czech':'čeština','Dutch':'nizozemština','Malay':'malajština','Hebrew':'hebrejština','Hindi':'hindi','TraditionalChinese':'tradiční čínština','Vietnamese':'vietnamština'},
 'de': {'Thai':'Thai','Indonesian':'Indonesisch','Russian':'Russisch','Arabic':'Arabisch','Portuguese':'Portugiesisch','Italian':'Italienisch','Polish':'Polnisch','Czech':'Tschechisch','Dutch':'Niederländisch','Malay':'Malaiisch','Hebrew':'Hebräisch','Hindi':'Hindi','TraditionalChinese':'Traditionelles Chinesisch','Vietnamese':'Vietnamesisch'},
 'en': {'Thai':'Thai','Indonesian':'Indonesian','Russian':'Russian','Arabic':'Arabic','Portuguese':'Portuguese','Italian':'Italian','Polish':'Polish','Czech':'Czech','Dutch':'Dutch','Malay':'Malay','Hebrew':'Hebrew','Hindi':'Hindi','TraditionalChinese':'Traditional Chinese','Vietnamese':'Vietnamese'},
 'es': {'Thai':'tailandés','Indonesian':'indonesio','Russian':'ruso','Arabic':'árabe','Portuguese':'portugués','Italian':'italiano','Polish':'polaco','Czech':'checo','Dutch':'neerlandés','Malay':'malayo','Hebrew':'hebreo','Hindi':'hindi','TraditionalChinese':'chino tradicional','Vietnamese':'vietnamita'},
 'fr': {'Thai':'thaï','Indonesian':'indonésien','Russian':'russe','Arabic':'arabe','Portuguese':'portuguais','Italian':'italien','Polish':'polonais','Czech':'tchèque','Dutch':'nederlands','Malay':'malais','Hebrew':'hébreu','Hindi':'hindi','TraditionalChinese':'chinois traditionnel','Vietnamese':'vietnamien'},
 'he': {'Thai':'תאית','Indonesian':'אינדונזית','Russian':'רוסית','Arabic':'ערבית','Portuguese':'פורטוגזית','Italian':'איטלקית','Polish':'פולנית','Czech':'צ׳כית','Dutch':'הולנדית','Malay':'מלאית','Hebrew':'עברית','Hindi':'הינדית','TraditionalChinese':'סינית מסורתית','Vietnamese':'וייטנאמית'},
 'hi': {'Thai':'थाई','Indonesian':'इंडोनेशियाई','Russian':'रूसी','Arabic':'अरबी','Portuguese':'पुर्तगाली','Italian':'इतालवी','Polish':'पोलिश','Czech':'चेक','Dutch':'डच','Malay':'मलय','Hebrew':'हिब्रू','Hindi':'हिंदी','TraditionalChinese':'पारंपरिक चीनी','Vietnamese':'वियतनामी'},
 'id': {'Thai':'Thai','Indonesian':'Bahasa Indonesia','Russian':'Rusia','Arabic':'Arab','Portuguese':'Portugis','Italian':'Italia','Polish':'Polandia','Czech':'Ceko','Dutch':'Belanda','Malay':'Melayu','Hebrew':'Ibrani','Hindi':'Hindi','TraditionalChinese':'Tionghoa Tradisional','Vietnamese':'Vietnam'},
 'it': {'Thai':'thai','Indonesian':'indonesiano','Russian':'russo','Arabic':'arabo','Portuguese':'portoghese','Italian':'italiano','Polish':'polacco','Czech':'ceco','Dutch':'olandese','Malay':'malese','Hebrew':'ebraico','Hindi':'hindi','TraditionalChinese':'cinese tradizionale','Vietnamese':'vietnamita'},
 'ja': {'Thai':'タイ語','Indonesian':'インドネシア語','Russian':'ロシア語','Arabic':'アラビア語','Portuguese':'ポルトガル語','Italian':'イタリア語','Polish':'ポーランド語','Czech':'チェコ語','Dutch':'オランダ語','Malay':'マレー語','Hebrew':'ヘブライ語','Hindi':'ヒンディー語','TraditionalChinese':'繁体字中国語','Vietnamese':'ベトナム語'},
 'ko': {'Thai':'태국어','Indonesian':'인도네시아어','Russian':'러시아어','Arabic':'아랍어','Portuguese':'포르투갈어','Italian':'이탈리아어','Polish':'폴란드어','Czech':'체코어','Dutch':'네덜란드어','Malay':'말레이어','Hebrew':'히브리어','Hindi':'힌디어','TraditionalChinese':'중국어 번체','Vietnamese':'베트남어'},
 'ms': {'Thai':'Thai','Indonesian':'Bahasa Indonesia','Russian':'Rusia','Arabic':'Arab','Portuguese':'Portugis','Italian':'Itali','Polish':'Poland','Czech':'Czech','Dutch':'Belanda','Malay':'Melayu','Hebrew':'Ibrani','Hindi':'Hindi','TraditionalChinese':'Cina Tradisional','Vietnamese':'Vietnam'},
 'nl': {'Thai':'Thai','Indonesian':'Indonesisch','Russian':'Russisch','Arabic':'Arabisch','Portuguese':'Portugees','Italian':'Italiaans','Polish':'Pools','Czech':'Tsjechisch','Dutch':'Nederlands','Malay':'Maleis','Hebrew':'Hebreeuws','Hindi':'Hindi','TraditionalChinese':'Traditioneel Chinees','Vietnamese':'Vietnamees'},
 'pl': {'Thai':'tajski','Indonesian':'indonejski','Russian':'rosyjski','Arabic':'arabski','Portuguese':'portugalski','Italian':'włoski','Polish':'polski','Czech':'czeski','Dutch':'niderlandzki','Malay':'malajski','Hebrew':'hebrajski','Hindi':'hindi','TraditionalChinese':'tradycyjny chiński','Vietnamese':'wietnamski'},
 'pt': {'Thai':'tailandês','Indonesian':'indonésio','Russian':'russo','Arabic':'árabe','Portuguese':'português','Italian':'italiano','Polish':'polaco','Czech':'tcheco','Dutch':'neerlandês','Malay':'malaio','Hebrew':'hebraico','Hindi':'híndi','TraditionalChinese':'chinês tradicional','Vietnamese':'vietnamita'},
 'ru': {'Thai':'тайский','Indonesian':'индонезийский','Russian':'русский','Arabic':'арабский','Portuguese':'португальский','Italian':'итальянский','Polish':'польский','Czech':'чешский','Dutch':'нидерландский','Malay':'малайский','Hebrew':'иврит','Hindi':'хинди','TraditionalChinese':'китайский (традиционный)','Vietnamese':'вьетнамский'},
 'th': {'Thai':'ภาษาไทย','Indonesian':'ภาษาอินโดนีเซีย','Russian':'ภาษารัสเซีย','Arabic':'ภาษาอาหรับ','Portuguese':'ภาษาโปรตุเกส','Italian':'ภาษาอิตาลี','Polish':'ภาษาโปแลนด์','Czech':'ภาษาเช็ก','Dutch':'ภาษาดัตช์','Malay':'ภาษามลายู','Hebrew':'ภาษาฮีบรู','Hindi':'ภาษาฮินดี','TraditionalChinese':'ภาษาจีนตัวเต็ม','Vietnamese':'ภาษาเวียดนาม'},
 'vi': {'Thai':'Tiếng Thái','Indonesian':'Tiếng Indonesia','Russian':'Tiếng Nga','Arabic':'Tiếng Ả Rập','Portuguese':'Tiếng Bồ Đào Nha','Italian':'Tiếng Ý','Polish':'Tiếng Ba Lan','Czech':'Tiếng Séc','Dutch':'Tiếng Hà Lan','Malay':'Tiếng Mã Lai','Hebrew':'Tiếng Do Thái','Hindi':'Tiếng Hindi','TraditionalChinese':'Tiếng Trung (Phồn thể)','Vietnamese':'Tiếng Việt'},
 'zh': {'Thai':'泰文','Indonesian':'印尼文','Russian':'俄文','Arabic':'阿拉伯文','Portuguese':'葡萄牙文','Italian':'意大利文','Polish':'波兰文','Czech':'捷克文','Dutch':'荷兰文','Malay':'马来文','Hebrew':'希伯来文','Hindi':'印地文','TraditionalChinese':'繁体中文','Vietnamese':'越南文'},
 'zh-TW': {'Thai':'泰文','Indonesian':'印尼文','Russian':'俄文','Arabic':'阿拉伯文','Portuguese':'葡萄牙文','Italian':'義大利文','Polish':'波蘭文','Czech':'捷克文','Dutch':'荷蘭文','Malay':'馬來文','Hebrew':'希伯來文','Hindi':'印地文','TraditionalChinese':'繁體中文','Vietnamese':'越南文'},
}

TREES = [
    ('docs app', ROOT / 'apps/docs/src/renderer/i18n/app', 'appLang'),
    ('docs ribbon', ROOT / 'apps/docs/src/renderer/i18n/ribbon', 'ribbonLang'),
    ('slides ribbon', ROOT / 'apps/slides/src/renderer/i18n/ribbon', 'ribbonLang'),
]


def patch(path: Path, family: str, locale: str) -> int:
    src = path.read_text(encoding='utf8')
    existing = set(re.findall(rf'\b{family}(\w+):', src))
    missing = [s for s in NEW if s not in existing]
    if not missing:
        return 0
    # insert right after the last existing family key, keeping the block together
    last = None
    for m in re.finditer(rf'^\s*{family}\w+:\s*.*$', src, re.M):
        last = m
    if last is None:
        raise SystemExit(f'no {family}* key found in {path}')
    lines = [f"  {family}{s}: '{T[locale][s]}'," for s in missing]
    src = src[:last.end()] + '\n' + '\n'.join(lines) + src[last.end():]
    path.write_text(src, encoding='utf8')
    return len(missing)


total = 0
for label, dirpath, family in TREES:
    n = 0
    for locale in T:
        n += patch(dirpath / f'{locale}.ts', family, locale)
    total += n
    print(f'{label:14} +{n} keys')
print(f'total +{total}')
