// @ts-nocheck — generation layer ported verbatim from untyped JS; it is typed
// file by file without logic changes, and until then strict consumers
// (apps/html, apps/shell) must not fail on it.
// Han + kana + Hangul: Korean previously fell outside this class and whole
// Korean documents took the Latin/Arial branch (serif Hangul fallback).
// CJK radicals/supplement, kana extensions, Hangul Jamo Extended-A/B and the
// astral CJK Extensions B-H (u-flag: for..of yields whole code points, so a
// BMP-only class never matches them) — without these, kana-ext/radical/Ext-B
// text falls through to the Latin/Arial branch.
const CJK_RE =
  /[\u1100-\u11ff\u2e80-\u2fdf\u3000-\u30ff\u31f0-\u31ff\u3130-\u318f\u3400-\u9fff\ua960-\ua97f\uac00-\ud7af\ud7b0-\ud7ff\uf900-\ufaff\uff00-\uffef\u{20000}-\u{3134f}]/u
const HANGUL_RE = /[\u1100-\u11ff\u3130-\u318f\uac00-\ud7af]/
const KANA_RE = /[\u3040-\u30ff\u31f0-\u31ff]/

const FONT_RULES: Array<[RegExp, string]> = [
  [/jetbrains|consolas|menlo|courier|mono/i, 'Consolas'],
  [/fredoka/i, 'Arial Rounded MT Bold'],
  [
    /brush script|lucida handwriting|pacifico|dancing script|great vibes|cursive|comic sans/i,
    'Segoe Script',
  ],
  [/dm serif|playfair|cormorant/i, 'Georgia'],
  [/archivo|bebas|oswald/i, 'Arial Narrow'],
  [/malgun|맑은/i, 'Malgun Gothic'],
  [/yu gothic|meiryo|hiragino kaku|ms gothic|noto sans jp/i, 'Yu Gothic'],
  [/batang|바탕|noto serif kr/i, 'Batang'],
  [/yu mincho|ms mincho|hiragino mincho|noto serif jp/i, 'Yu Mincho'],
  [
    /noto sans cjk|noto sans sc|noto sans tc|noto sans kr|yahei|pingfang|heiti|source han sans|\u5fae\u8f6f\u96c5\u9ed1|\u9ed1\u4f53/i,
    'Microsoft YaHei',
  ],
  [
    /noto serif cjk|noto serif sc|songti|stsong|simsun|kaiti|fangsong|source han serif|\u5b8b\u4f53/i,
    'SimSun',
  ],
  [/arial unicode/i, 'Arial Unicode MS'],
  [/arial|helvetica|verdana|tahoma|segoe|roboto|inter|lato|open sans|sans/i, 'Arial'],
  [/times|georgia|garamond|cambria|palatino|book antiqua|serif/i, 'Times New Roman'],
  [/.*/, 'Arial'],
]

function mapFont(fontFamily = '', mono = false, text = '') {
  if (mono) return 'Consolas'
  let font = 'Arial'
  for (const [pattern, name] of FONT_RULES) {
    if (pattern.test(fontFamily)) {
      font = name
      break
    }
  }
  if (CJK_RE.test(text)) {
    const sans = HANGUL_RE.test(text)
      ? 'Malgun Gothic'
      : KANA_RE.test(text)
        ? 'Yu Gothic'
        : 'Microsoft YaHei'
    const serif = HANGUL_RE.test(text) ? 'Batang' : KANA_RE.test(text) ? 'Yu Mincho' : 'SimSun'
    if (font === 'Arial') font = sans
    if (font === 'Times New Roman') font = serif
  }
  return font
}

const CJK_SANS = new Set(['Microsoft YaHei', 'Malgun Gothic', 'Yu Gothic'])
const CJK_SERIF = new Set(['SimSun', 'Batang', 'Yu Mincho'])

/** Run font for docx: a CJK face goes on eastAsia only, since Word draws
 *  Latin with a CJK font's half-width glyphs (the SimSun monospace look). */
function runFont(fontFamily = '', mono = false, text = '') {
  const font = mapFont(fontFamily, mono, text)
  if (CJK_SANS.has(font)) return { ascii: 'Arial', hAnsi: 'Arial', cs: 'Arial', eastAsia: font }
  if (CJK_SERIF.has(font))
    return {
      ascii: 'Times New Roman',
      hAnsi: 'Times New Roman',
      cs: 'Times New Roman',
      eastAsia: font,
    }
  return font
}

export { CJK_RE, HANGUL_RE, KANA_RE, FONT_RULES, mapFont, runFont }
