import type { Lang } from '@genoffice/i18n'

/**
 * Font dropdown candidates grouped by script, ordered per UI language so the
 * fonts a user is most likely to want appear at the top (e.g. Western fonts
 * for the US market, Japanese fonts for the Japanese market).
 */
const LATIN = [
  'Calibri',
  'Arial',
  'Times New Roman',
  'Georgia',
  'Verdana',
  'Tahoma',
  'Cambria',
  'Garamond',
  'Trebuchet MS',
  'Segoe UI',
  'Courier New',
  'Impact',
]

// Bundled Google Fonts (fonts/google-fonts.css, latin subset woff2) — a
// "Google Fonts" group under the Word-standard set above. font-list.ts has
// no grouping primitive (fontFamiliesFor returns one flat array per script),
// so this is a plain array appended after LATIN and the toolbar renders it
// merged into the same list, same as every other entry.
const GOOGLE_FONTS = [
  'Roboto',
  'Open Sans',
  'Lato',
  'Montserrat',
  'Inter',
  'Work Sans',
  'Poppins',
  'Oswald',
  'Raleway',
  'Nunito',
  'Merriweather',
  'Playfair Display',
  'Lora',
  'Source Serif 4',
  'Source Sans 3',
  'PT Serif',
  'PT Sans',
  'Libre Franklin',
  'EB Garamond',
  'Noto Serif',
  'Rubik',
  'Karla',
  'DM Sans',
  'Space Grotesk',
  'Bebas Neue',
  'Archivo',
]
// GB/T 9704 official-document fonts included so government documents can be
// authored from scratch (values are free-typed either way; the combobox accepts any name)
const SIMPLIFIED_CHINESE = [
  '等线',
  '宋体',
  '黑体',
  '微软雅黑',
  '楷体',
  '仿宋',
  '仿宋_GB2312',
  '楷体_GB2312',
  '方正小标宋简体',
]
const JAPANESE = ['Yu Gothic', 'Yu Mincho', 'Meiryo', 'MS Gothic', 'MS Mincho']
const KOREAN = ['Malgun Gothic', 'Batang', 'Gulim', 'Dotum']
const TRADITIONAL_CHINESE = ['Microsoft JhengHei', 'PMingLiU']

export const BUILTIN_FONT_FAMILIES: readonly string[] = [
  ...LATIN,
  ...SIMPLIFIED_CHINESE,
  ...JAPANESE,
  ...KOREAN,
  ...TRADITIONAL_CHINESE,
]

const EAST_ASIAN_FONT_RE =
  /[⺀-鿿豈-﫿぀-ヿㇰ-ㇿ가-힯]|sim(sun|hei)|nsimsun|kaiti|fangsong|dengxian|yahei|songti|heiti|xingkai|lisu|youyuan|st(zhongsong|song|kai|fangsong|xihei|hupo|liti|caiyun)|pingfang|hiragino|meiryo|osaka|kozuka|yu (gothic|mincho)|yugoth|ms (ui )?p?(gothic|mincho)|biz ud|malgun|batang|gulim|dotum|gungsuh|m(ye|yu)ngjo|nanum|apple (sd )?gothic|applemyungjo|jhenghei|p?mingliu|biaukai|dfkai|kaiu|source han|noto (sans|serif) (cjk|sc|tc|hk|jp|kr)|wenquanyi/i

/**
 * Which rFonts slot a font-box pick should target: East Asian names go to
 * w:eastAsia, everything else to w:ascii/w:hAnsi — mirroring Word, where
 * picking a Latin font never clobbers the Chinese font and vice versa.
 */
export function isEastAsianFontName(name: string): boolean {
  return EAST_ASIAN_FONT_RE.test(name.normalize('NFKC'))
}

export function fontFamiliesFor(lang: Lang): readonly string[] {
  switch (lang) {
    case 'zh':
      return [
        ...SIMPLIFIED_CHINESE,
        ...LATIN,
        ...GOOGLE_FONTS,
        ...JAPANESE,
        ...KOREAN,
        ...TRADITIONAL_CHINESE,
      ]
    case 'zh-TW':
      return [
        ...TRADITIONAL_CHINESE,
        ...LATIN,
        ...GOOGLE_FONTS,
        ...SIMPLIFIED_CHINESE,
        ...JAPANESE,
        ...KOREAN,
      ]
    case 'ja':
      return [
        ...JAPANESE,
        ...LATIN,
        ...GOOGLE_FONTS,
        ...SIMPLIFIED_CHINESE,
        ...KOREAN,
        ...TRADITIONAL_CHINESE,
      ]
    case 'ko':
      return [
        ...KOREAN,
        ...LATIN,
        ...GOOGLE_FONTS,
        ...JAPANESE,
        ...SIMPLIFIED_CHINESE,
        ...TRADITIONAL_CHINESE,
      ]
    default:
      return [
        ...LATIN,
        ...GOOGLE_FONTS,
        ...JAPANESE,
        ...SIMPLIFIED_CHINESE,
        ...KOREAN,
        ...TRADITIONAL_CHINESE,
      ]
  }
}

/**
 * docDefaults w:eastAsia font for new blank documents, matching what Word
 * ships per market (zh → SimSun, ja → Yu Mincho, ko → Malgun Gothic,
 * zh-TW → PMingLiU). English and every other language return undefined —
 * like en-US Word, whose theme leaves the East Asian slot empty and lets
 * per-script substitution kick in only when CJK text actually appears.
 * Latin default is Work Sans for every language (see blank.ts stylesXml).
 */
export function defaultEastAsiaFontFor(lang: Lang): string | undefined {
  switch (lang) {
    case 'zh':
      return '宋体'
    case 'ja':
      return 'Yu Mincho'
    case 'ko':
      return 'Malgun Gothic'
    case 'zh-TW':
      return 'PMingLiU'
    default:
      return undefined
  }
}
