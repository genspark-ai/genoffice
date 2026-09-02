/**
 * Cell-font fallback for the grid canvas (document data, theme-independent).
 * Univer draws/measures with a bare `italic bold 11pt "Family"` string, so an
 * unresolvable family drops to Chromium's default *serif* chain while Excel's
 * substitution is always sans. Two layers: a ctx.font setter patch appends
 * `sans-serif` as the last resort, and alias FontFaces map well-known Office
 * names to local faces — pinning serif-intent names (Mincho/Song/Ming/Batang/
 * Cambria…) to real serif faces so the sans last-resort never flips them.
 * Where no alias local() face exists at all (e.g. Linux), alias-known serif
 * names still get `serif` appended instead of `sans-serif`; unrecognized
 * names always fall to sans, matching Excel's substitution.
 */

const GENERIC_FAMILY =
  /(?:^|[\s,])(?:serif|sans-serif|monospace|cursive|fantasy|system-ui|math|ui-serif|ui-sans-serif|ui-monospace|ui-rounded)$/i

const SERIF_INTENT =
  /mincho|明朝|simsun|songti|宋体|宋體|mingliu|明體|細明|batang|바탕|myeongjo|명조|roman|georgia|garamond|cambria|constantia|palatino|antiqua|didot|bodoni|baskerville|caslon|goudy|bookman|(?<!sans[-\s])serif/i

/// Generic families never cover emoji code points on the canvas — without an
/// explicit color-emoji face at the end of every chain, U+274C & friends draw
/// as tofu. Listed after the generic so primary-font metrics never change.
/// Emoji=Yes but Emoji_Presentation=No dingbats (✔✖❄❤ card suits …) render
/// as text glyphs in the cell color in Excel, while a color-emoji bitmap
/// ignores fillStyle (prod_016's CF-red ✖ drew charcoal) — a monochrome
/// symbols face intercepts exactly those codepoints ahead of the emoji
/// chain; EPres=Yes codepoints (U+274C …) stay on the color font.
const TEXT_DINGBATS_FAMILY = 'Cell Text Dingbats'
/// Faces that carry real outline glyphs for these codepoints (canvas-probed:
/// Apple Symbols and Hiragino draw the color-emoji bitmap instead).
const TEXT_DINGBATS_SOURCES = [
  'Segoe UI Symbol',
  'Arial Unicode MS',
  'Zapf Dingbats',
  'Menlo',
] as const
const TEXT_DINGBATS_RANGE =
  'U+2611, U+2660, U+2663, U+2665-2666, U+2702, U+2708-2709, U+270C, U+270F, ' +
  'U+2712, U+2714, U+2716, U+271D, U+2721, U+2733-2734, U+2744, U+2747, U+2763-2764'
const EMOJI_FALLBACK =
  '"Cell Text Dingbats", "Apple Color Emoji", "Segoe UI Emoji", "Noto Color Emoji"'

/// The family name inside one comma segment of a ctx.font string: the quoted
/// name if quoted, otherwise the last whitespace token (the first segment
/// carries the style/size prefix; Univer quotes any family with spaces).
function segmentFamily(segment: string): string {
  const quoted = /["']([^"']*)["']/.exec(segment)
  if (quoted) return quoted[1]!.trim()
  const tokens = segment.trim().split(/\s+/)
  return tokens[tokens.length - 1] ?? ''
}

/// Only alias-known names carry serif intent (first hit in the list wins).
/// Excel substitutes its sans default for any name it cannot resolve — even
/// myeongjo/mincho-keyworded ones ("휴먼명조,한컴돋움" renders sans in the
/// Excel reference for prod_059, and Univer's per-glyph fallback re-probes
/// each member alone, so the single-name path must agree). Alias-known serif
/// names keep serif so the last resort never flips them where the alias
/// local() faces are absent (e.g. Linux).
function serifIntent(trimmed: string): boolean {
  for (const segment of trimmed.split(',')) {
    const family = segmentFamily(segment)
    if (ALIAS_FAMILY_NAMES.has(family.toLowerCase())) return SERIF_INTENT.test(family)
  }
  return false
}

export function withSansSerifFallback(font: string): string {
  const trimmed = font.trimEnd()
  if (!trimmed) return font
  if (trimmed.endsWith(EMOJI_FALLBACK)) return font
  if (GENERIC_FAMILY.test(trimmed)) return `${trimmed}, ${EMOJI_FALLBACK}`
  return `${trimmed}, ${serifIntent(trimmed) ? 'serif' : 'sans-serif'}, ${EMOJI_FALLBACK}`
}

const canvasScopedFamilies = new Map<string, string>()
/// Families whose registered face is a substitute (the skipIfLocal probe
/// found no genuine font): only these carry alias-calibrated metrics.
const substitutedFamilies = new Set<string>()

export function isSubstitutedCellFamily(family: string): boolean {
  return substitutedFamilies.has(family.toLowerCase())
}

const PASSTHROUGH_FAMILY =
  /^(?:serif|sans-serif|monospace|cursive|fantasy|system-ui|math|ui-serif|ui-sans-serif|ui-monospace|ui-rounded|-apple-system|BlinkMacSystemFont)$/i

export function rewriteScopedFamilies(
  font: string,
  scoped: ReadonlyMap<string, string> = canvasScopedFamilies,
): string {
  if (scoped.size === 0) return font
  const segments = font.split(',')
  const families = segments.map((s) => segmentFamily(s))
  const scopedIdx = families.findIndex((f) => scoped.has(f.toLowerCase()))
  if (scopedIdx === -1) return font
  // Substitute only a sole cell family (plus generics): an explicit fallback
  // stack is a UI measurement mirroring CSS and must fall through natively,
  // exactly like the DOM it matches.
  if (families.some((f, i) => i !== scopedIdx && !PASSTHROUGH_FAMILY.test(f))) return font
  const target = scoped.get(families[scopedIdx]!.toLowerCase())!
  const segment = segments[scopedIdx]!
  const quoted = /["'][^"']*["']/.exec(segment)
  segments[scopedIdx] = quoted
    ? segment.replace(quoted[0], `"${target}"`)
    : segment.replace(/\S+(\s*)$/, `"${target}"$1`)
  return segments.join(',')
}

export interface CellFontAlias {
  readonly family: string
  /// local() face names, tried in order; genuine (Windows) names first so the
  /// alias is a no-op where the real font exists. Items containing '(' are
  /// raw src tokens (e.g. url(...) for bundled fonts).
  readonly regular: readonly string[]
  /// Real bold faces only — never a regular face, which would suppress
  /// synthetic bold where no true bold exists.
  readonly bold?: readonly string[]
  /// size-adjust % matching the substitute's advances to the original's
  /// Excel-print advances (weighted per-char ratio measured from production
  /// ref PDFs). Requires skipIfLocal: the adjustment is derived for the
  /// substitute face and must never distort the genuine font.
  readonly sizeAdjust?: string
  readonly boldSizeAdjust?: string
  /// local() names (family/PostScript/full) proving the genuine font exists;
  /// when any resolves, the whole alias is skipped so real metrics win.
  readonly skipIfLocal?: readonly string[]
  /// Faces registered instead when skipIfLocal finds the genuine font: a
  /// plain rename mapping (no width correction) for spellings the OS matcher
  /// cannot resolve itself — Chromium on macOS never matches localized
  /// family names, so '맑은 고딕' needs an explicit map to Malgun Gothic.
  readonly whenGenuine?: { readonly regular: readonly string[]; readonly bold?: readonly string[] }
  /// Register the faces under an internal name reachable only through the
  /// patched ctx.font setter: UI chrome stacks reference 'Segoe UI', and a
  /// document-wide web face under that name would restyle the ribbon.
  readonly scopeToCanvas?: true
  /// Latin/digit sub-face registered after the base faces with
  /// unicodeRange U+0-2CFF, so dual-metric fonts (e.g. Malgun Gothic:
  /// hangul matches AppleGothic exactly, digits do not) can correct each
  /// script independently.
  readonly latin?: {
    readonly regular: readonly string[]
    readonly bold?: readonly string[]
    readonly sizeAdjust?: string
    readonly boldSizeAdjust?: string
    /// Overrides the default U+0-2CFF span — Thai sits inside it, so Thai
    /// aliases carve their own block out and keep it on the base face.
    readonly unicodeRange?: string
  }
}

const JP_SANS = ['Hiragino Sans', 'HiraginoSans-W3', 'Hiragino Kaku Gothic ProN'] as const
const JP_SANS_BOLD = ['HiraginoSans-W6', 'Hiragino Sans W6'] as const
const JP_SERIF = ['Hiragino Mincho ProN', 'HiraMinProN-W3'] as const
const JP_SERIF_BOLD = ['HiraMinProN-W6', 'Hiragino Mincho ProN W6'] as const
const YU_GOTHIC = ['Yu Gothic Regular', 'YuGothic-Regular', 'YuGothic Medium', 'YuGo-Medium']
const YU_GOTHIC_BOLD = ['Yu Gothic Bold', 'YuGothic-Bold', 'YuGothic Bold', 'YuGo-Bold']
const YU_MINCHO = ['Yu Mincho Regular', 'YuMincho-Regular', 'YuMincho Medium', 'YuMin-Medium']
const YU_MINCHO_BOLD = ['Yu Mincho Demibold', 'YuMincho-Demibold', 'YuMin-Demibold']
const YU_MINCHO_ALL_BOLD = [...YU_MINCHO_BOLD, ...JP_SERIF_BOLD]
const SONG = ['SimSun', 'Songti SC', 'STSongti-SC-Regular']
const SONG_BOLD = ['STSongti-SC-Bold', 'Songti SC Bold']
const KAI = ['KaiTi', 'Kaiti SC', 'STKaitiSC-Regular', 'STKaiti']
const MING_TC = ['PMingLiU', 'Songti TC', 'Apple LiSung']
const KR_SANS = ['Malgun Gothic', 'Apple SD Gothic Neo', 'AppleGothic']
/// Carlito is bundled, not installed — local() alone can never resolve it.
const CARLITO_SRC = [
  'Carlito',
  `url(${new URL('./fonts/Carlito-Regular.ttf', import.meta.url).href})`,
]
const CARLITO_BOLD_SRC = [
  'Carlito Bold',
  `url(${new URL('./fonts/Carlito-Bold.ttf', import.meta.url).href})`,
]
/// Malgun Gothic prints hangul at 1.0em — exactly AppleGothic — but digits at
/// 0.6em vs AppleGothic's 0.68em, so number tails clipped while hangul was
/// perfect. Latin/digit runs go to width-corrected Helvetica Neue instead.
const MALGUN_ALIAS: Omit<CellFontAlias, 'family'> = {
  regular: ['Malgun Gothic', 'AppleGothic'],
  skipIfLocal: ['Malgun Gothic', 'MalgunGothic'],
  latin: {
    regular: ['Helvetica Neue'],
    sizeAdjust: '104%',
    bold: ['Helvetica Neue Bold'],
    boldSizeAdjust: '109.4%',
  },
}
const KR_SERIF = ['Batang', 'AppleMyungjo', 'Nanum Myeongjo']
const TIMES_BOLD = ['Times New Roman Bold', 'TimesNewRomanPS-BoldMT']

type LatinSubFace = NonNullable<CellFontAlias['latin']>

/// Width-corrected Latin/digit sub-face on Helvetica Neue (the design closest
/// to the Office sans faces' Latin, and the face the other aliases use).
function helveticaLatin(sizeAdjust: string, boldSizeAdjust: string): LatinSubFace {
  return {
    regular: ['Helvetica Neue'],
    sizeAdjust,
    bold: ['Helvetica Neue Bold'],
    boldSizeAdjust,
  }
}

// JP gothic substitutes. Kanji are 1.0em on both the Windows faces and
// Hiragino Sans, so the base face needs no correction, but the Windows faces
// keep much narrower Latin: MS (P/UI) Gothic digits are 0.5em (JIS
// half-width) and Yu Gothic's 0.556em against Hiragino Sans' 0.657em, so
// Latin/digit tails clipped in columns Excel had fitted. Ratios are
// hmtx-exact (msgothic.ttc / YuGoth*.ttc / meiryo.ttc as shipped with Mac
// Office) weighted by the Latin character mix of the ja prod refs (prod_022 /
// prod_037 / prod_047 / prod_038 MS PGothic runs, 2191 non-space chars).
// Known limit: MS PGothic / MS UI Gothic / Meiryo UI kana are proportional
// (0.74-0.90em) while Hiragino kana stay 1.0em — kana-heavy strings remain
// wider than Excel's; only the Latin range is corrected here.
/// MS PGothic / MS UI Gothic: 0.5em digits vs Helvetica Neue 0.556em. The
/// family has no bold face and Excel's synthetic bold keeps the regular
/// advances, so the bold sub-face targets the same widths.
const MS_GOTHIC_LATIN = helveticaLatin('90.9%', '90.1%')
/// MS Gothic's Latin is half-width monospace (every glyph 0.5em); the
/// proportional substitute pins digits — the alignment-critical class — at
/// exactly 0.5em and lets letters average out around it.
const MS_GOTHIC_MONO_LATIN = helveticaLatin('89.9%', '89.9%')
/// Yu Gothic's Latin is metrically a Helvetica Neue clone (digits 0.5562em vs
/// 0.556em); the bold face runs 3% wider.
const YU_GOTHIC_LATIN = helveticaLatin('100.8%', '103.3%')
const YU_GOTHIC_UI_LATIN = helveticaLatin('95.7%', '94.6%')
/// Meiryo's Latin is the Verdana design (digits 0.621em vs Verdana 0.636em,
/// uniform 97-98% across digits / capitals / lowercase), so Verdana keeps its
/// look as well as its widths.
const MEIRYO_LATIN: LatinSubFace = {
  regular: ['Verdana'],
  sizeAdjust: '97.6%',
  bold: ['Verdana Bold', 'Verdana-Bold'],
  boldSizeAdjust: '95.2%',
}

/// Hiragino base for kana/kanji plus a corrected Latin sub-face, registered
/// only where the genuine face is absent; `genuine` doubles as the plain
/// rename map (no correction) when it exists, since Chromium on macOS never
/// matches localized family names ('ＭＳ Ｐゴシック' → 'MS PGothic').
function jpGothic(
  family: string,
  genuine: readonly string[],
  latin: LatinSubFace,
  genuineBold?: readonly string[],
): CellFontAlias {
  return {
    family,
    regular: [...JP_SANS],
    bold: [...JP_SANS_BOLD],
    skipIfLocal: genuine,
    whenGenuine: genuineBold ? { regular: genuine, bold: genuineBold } : { regular: genuine },
    latin,
  }
}

function jpMincho(family: string, genuine: readonly string[]): CellFontAlias {
  return { family, regular: [...genuine, ...JP_SERIF], bold: [...JP_SERIF_BOLD] }
}

// Thai Office faces (Cordia New / Angsana New / TH Sarabun; the UPC spellings
// are the same designs under the Thai-codepage names) are absent on macOS.
// Their glyphs sit very small in the em box — Cordia New digits are 0.3645em
// and its Thai consonants ~0.30em weighted — while Thonburi, the macOS Thai
// face, draws Thai at ~0.44em and digits at 0.666em, so table headers Excel
// had fitted clipped by a third (prod_066). Thai glyphs go to Thonburi and
// the Latin/digit range to a Latin face, each with its own size-adjust; Thai
// combining marks are zero-width on both sides. Ratios are hmtx-exact from
// the font files Mac Office ships app-private (cordia.ttc / angsa.ttf /
// thsarabun.ttf) weighted by prod_066's character mix (3376 Thai / 7434
// Latin chars, 4531 of them digits), and agree with the glyph positions of
// its Excel reference print within 0.4%.
const THAI_SANS = ['Thonburi']
const THAI_SANS_BOLD = ['Thonburi-Bold', 'Thonburi Bold']
/// Thai aliases keep U+0E00-0E7F on the base face.
const LATIN_RANGE_EXCLUDING_THAI = 'U+0-DFF, U+E80-2CFF'

interface ThaiAliasMetrics {
  /// Thonburi size-adjust for the Thai block (regular / bold).
  readonly thai: string
  readonly thaiBold: string
  readonly latin: LatinSubFace
}

function thaiAlias(
  family: string,
  skipIfLocal: readonly string[],
  metrics: ThaiAliasMetrics,
): CellFontAlias {
  return {
    family,
    regular: THAI_SANS,
    sizeAdjust: metrics.thai,
    bold: THAI_SANS_BOLD,
    boldSizeAdjust: metrics.thaiBold,
    skipIfLocal,
    latin: { ...metrics.latin, unicodeRange: LATIN_RANGE_EXCLUDING_THAI },
  }
}

/// Cordia New: Thai 0.2975em (bold 0.3505em) vs Thonburi 0.4423em; Latin
/// 0.3267em vs Helvetica Neue 0.5029em. Helvetica Neue at 65% also lands
/// Cordia's cap height (0.714 × 0.65 = 0.464em vs 0.469em).
const CORDIA_METRICS: ThaiAliasMetrics = {
  thai: '67.3%',
  thaiBold: '74.1%',
  latin: helveticaLatin('65%', '66%'),
}
/// Angsana New's Latin is Times New Roman scaled to 66% — every class
/// (digits / capitals / lowercase) gives the same ratio — so it keeps its
/// serif look; Thai 0.2972em vs Thonburi 0.4449em.
const ANGSANA_METRICS: ThaiAliasMetrics = {
  thai: '66.8%',
  thaiBold: '66.9%',
  latin: {
    regular: ['Times New Roman'],
    sizeAdjust: '66%',
    bold: TIMES_BOLD,
    boldSizeAdjust: '66%',
  },
}
/// TH SarabunPSK: Thai 0.2882em (bold 0.3007em) vs Thonburi 0.4449em; Latin
/// 0.3252em (bold 0.3406em) vs Helvetica Neue 0.5013em (bold 0.5068em). TH
/// Sarabun New is the SIPA re-release of the same design with the same
/// advance widths.
const SARABUN_METRICS: ThaiAliasMetrics = {
  thai: '64.8%',
  thaiBold: '67.6%',
  latin: helveticaLatin('64.9%', '67.2%'),
}

export const CELL_FONT_ALIASES: readonly CellFontAlias[] = [
  // JP gothic (sans intent), incl. fullwidth spellings
  jpGothic('ＭＳ Ｐゴシック', ['MS PGothic', 'MS-PGothic'], MS_GOTHIC_LATIN),
  jpGothic('ＭＳ ゴシック', ['MS Gothic', 'MS-Gothic'], MS_GOTHIC_MONO_LATIN),
  jpGothic('MS PGothic', ['MS PGothic', 'MS-PGothic'], MS_GOTHIC_LATIN),
  jpGothic('MS Gothic', ['MS Gothic', 'MS-Gothic'], MS_GOTHIC_MONO_LATIN),
  jpGothic('MS UI Gothic', ['MS UI Gothic', 'MS-UIGothic'], MS_GOTHIC_LATIN),
  jpGothic('メイリオ', ['Meiryo'], MEIRYO_LATIN, ['Meiryo Bold', 'Meiryo-Bold']),
  jpGothic('Meiryo', ['Meiryo'], MEIRYO_LATIN, ['Meiryo Bold', 'Meiryo-Bold']),
  jpGothic('Meiryo UI', ['Meiryo UI', 'MeiryoUI'], MEIRYO_LATIN, [
    'Meiryo UI Bold',
    'MeiryoUI-Bold',
  ]),
  // Yu Gothic: the genuine gate also accepts the optional macOS YuGothic
  // faces (Medium stands in for Regular, as before).
  jpGothic('游ゴシック', ['Yu Gothic', ...YU_GOTHIC], YU_GOTHIC_LATIN, YU_GOTHIC_BOLD),
  jpGothic('游ゴシック体', ['Yu Gothic', ...YU_GOTHIC], YU_GOTHIC_LATIN, YU_GOTHIC_BOLD),
  jpGothic('Yu Gothic', ['Yu Gothic', ...YU_GOTHIC], YU_GOTHIC_LATIN, YU_GOTHIC_BOLD),
  jpGothic(
    'Yu Gothic UI',
    ['Yu Gothic UI', 'YuGothicUI-Regular', ...YU_GOTHIC],
    YU_GOTHIC_UI_LATIN,
    ['Yu Gothic UI Bold', 'YuGothicUI-Bold', ...YU_GOTHIC_BOLD],
  ),
  // JP mincho (serif intent — keep serif under the sans last-resort)
  jpMincho('ＭＳ 明朝', ['MS Mincho']),
  jpMincho('ＭＳ Ｐ明朝', ['MS PMincho']),
  jpMincho('MS Mincho', ['MS Mincho']),
  jpMincho('MS PMincho', ['MS PMincho']),
  { family: '游明朝', regular: [...YU_MINCHO, ...JP_SERIF], bold: YU_MINCHO_ALL_BOLD },
  { family: 'Yu Mincho', regular: [...YU_MINCHO, ...JP_SERIF], bold: YU_MINCHO_ALL_BOLD },
  // Simplified CJK
  { family: '宋体', regular: SONG, bold: SONG_BOLD },
  { family: 'SimSun', regular: SONG, bold: SONG_BOLD },
  { family: 'NSimSun', regular: SONG, bold: SONG_BOLD },
  { family: '黑体', regular: ['SimHei', 'Heiti SC', 'PingFang SC'] },
  { family: 'SimHei', regular: ['SimHei', 'Heiti SC', 'PingFang SC'] },
  { family: '仿宋', regular: ['FangSong', 'STFangsong'] },
  { family: 'FangSong', regular: ['FangSong', 'STFangsong'] },
  { family: '楷体', regular: KAI },
  { family: 'KaiTi', regular: KAI },
  { family: '楷体_GB2312', regular: KAI },
  // Traditional CJK
  { family: '新細明體', regular: MING_TC },
  { family: 'PMingLiU', regular: MING_TC },
  { family: '細明體', regular: ['MingLiU', ...MING_TC] },
  { family: 'MingLiU', regular: ['MingLiU', ...MING_TC] },
  { family: '標楷體', regular: ['DFKai-SB', 'BiauKai', 'Kaiti TC', 'Kaiti SC'] },
  { family: 'DFKai-SB', regular: ['DFKai-SB', 'BiauKai', 'Kaiti TC', 'Kaiti SC'] },
  // Korean
  { family: 'Malgun Gothic', ...MALGUN_ALIAS },
  {
    family: '맑은 고딕',
    ...MALGUN_ALIAS,
    whenGenuine: {
      regular: ['Malgun Gothic'],
      bold: ['Malgun Gothic Bold', 'MalgunGothicBold'],
    },
  },
  {
    family: '돋움',
    regular: ['Dotum', 'AppleGothic', 'Apple SD Gothic Neo'],
    bold: ['Dotum Bold', 'Apple SD Gothic Neo Bold', 'AppleSDGothicNeo-Bold'],
  },
  { family: 'Gulim', regular: ['Gulim', 'Dotum', ...KR_SANS] },
  { family: '굴림', regular: ['Gulim', 'Dotum', ...KR_SANS] },
  { family: 'Dotum', regular: ['Dotum', ...KR_SANS] },
  { family: 'Batang', regular: KR_SERIF },
  { family: '바탕', regular: KR_SERIF },
  { family: 'Gungsuh', regular: ['Gungsuh', 'AppleMyungjo'] },
  { family: '궁서', regular: ['Gungsuh', 'AppleMyungjo'] },
  // Office Latin serif faces absent on macOS (serif intent)
  {
    family: 'Cambria',
    regular: ['Cambria', 'Times New Roman', 'Georgia'],
    bold: ['Cambria Bold', ...TIMES_BOLD],
  },
  { family: 'Constantia', regular: ['Constantia', 'Georgia', 'Times New Roman'] },
  {
    family: 'Garamond',
    regular: ['Garamond', 'Times New Roman'],
    bold: ['Garamond Bold', ...TIMES_BOLD],
  },
  { family: 'Palatino Linotype', regular: ['Palatino Linotype', 'Palatino', 'Book Antiqua'] },
  { family: 'Book Antiqua', regular: ['Book Antiqua', 'Palatino'] },
  {
    family: 'Times New Roman',
    regular: ['Times New Roman', 'Times', 'Georgia'],
    bold: TIMES_BOLD,
  },
  { family: 'PT Serif', regular: ['PT Serif', 'Times New Roman', 'Georgia'] },
  // Width-corrected substitutes for fonts absent on macOS. Excel sized the
  // author's columns for the original font; a substitute with different
  // advances clips tail characters or wraps an extra line. size-adjust values
  // are weighted per-char advance ratios (original from production ref-PDF
  // glyph positions / substitute from live canvas measurement).
  {
    family: 'Bahnschrift',
    regular: ['Helvetica Neue'],
    sizeAdjust: '96.7%',
    bold: ['Helvetica Neue Bold'],
    boldSizeAdjust: '92.7%',
    skipIfLocal: ['Bahnschrift'],
  },
  {
    family: 'Segoe UI',
    regular: ['Helvetica Neue'],
    sizeAdjust: '96.7%',
    bold: ['Helvetica Neue Bold'],
    boldSizeAdjust: '98.3%',
    skipIfLocal: ['Segoe UI', 'SegoeUI'],
    scopeToCanvas: true,
  },
  {
    family: 'Dosis',
    regular: CARLITO_SRC,
    sizeAdjust: '96.3%',
    bold: CARLITO_BOLD_SRC,
    boldSizeAdjust: '99.2%',
    skipIfLocal: ['Dosis', 'Dosis-Regular', 'Dosis Regular'],
  },
  {
    family: 'Aptos Narrow',
    regular: CARLITO_SRC,
    sizeAdjust: '96%',
    bold: CARLITO_BOLD_SRC,
    boldSizeAdjust: '96%',
    skipIfLocal: ['Aptos Narrow', 'AptosNarrow'],
  },
  // Excel maps the Demi/Light family names onto the base family's bold and
  // regular when they are missing, so the substitutes mirror that weight
  // mapping; Helvetica Neue already matches LT Pro widths within 0.5%.
  {
    family: 'Avenir Next LT Pro',
    regular: ['AvenirNextLTPro-Regular', 'Helvetica Neue'],
    bold: ['AvenirNextLTPro-Bold', 'Helvetica Neue Bold'],
  },
  {
    family: 'Avenir Next LT Pro Demi',
    regular: ['AvenirNextLTPro-Demi', 'Helvetica Neue Bold'],
  },
  {
    family: 'Avenir Next LT Pro Light',
    regular: ['AvenirNextLTPro-Lt', 'Helvetica Neue'],
  },
  // Thai Office faces (see THAI_SANS above)
  thaiAlias('Cordia New', ['Cordia New', 'CordiaNew'], CORDIA_METRICS),
  thaiAlias('CordiaUPC', ['CordiaUPC'], CORDIA_METRICS),
  thaiAlias('Angsana New', ['Angsana New', 'AngsanaNew'], ANGSANA_METRICS),
  thaiAlias('AngsanaUPC', ['AngsanaUPC'], ANGSANA_METRICS),
  thaiAlias('TH SarabunPSK', ['TH SarabunPSK', 'THSarabunPSK'], SARABUN_METRICS),
  thaiAlias('TH Sarabun New', ['TH Sarabun New', 'THSarabunNew'], SARABUN_METRICS),
]

const ALIAS_FAMILY_NAMES: ReadonlySet<string> = new Set(
  CELL_FONT_ALIASES.map((alias) => alias.family.toLowerCase()),
)

function patchFontSetter(proto: object): void {
  const desc = Object.getOwnPropertyDescriptor(proto, 'font')
  if (!desc?.set || !desc.configurable) return
  const nativeSet = desc.set
  Object.defineProperty(proto, 'font', {
    ...desc,
    set(value: string) {
      nativeSet.call(this, withSansSerifFallback(rewriteScopedFamilies(String(value))))
    },
  })
}

export function installCanvasFontFallback(): void {
  patchFontSetter(CanvasRenderingContext2D.prototype)
  if (typeof OffscreenCanvasRenderingContext2D !== 'undefined')
    patchFontSetter(OffscreenCanvasRenderingContext2D.prototype)
}

/// Everything below the CJK blocks: latin sub-faces cover digits/latin/punct
/// while hangul & friends stay on the base face (a later-registered face wins
/// where unicode ranges overlap).
const LATIN_RANGE = 'U+0-2CFF'

function faceSrc(items: readonly string[]): string {
  return items.map((n) => (n.includes('(') ? n : `local('${n}')`)).join(', ')
}

function genuineLocalExists(names: readonly string[]): Promise<boolean> {
  const probes = names.map((n) =>
    new FontFace('__genoffice-font-probe', `local('${n}')`).load().then(
      () => true,
      () => false,
    ),
  )
  return Promise.all(probes).then((hits) => hits.some(Boolean))
}

/// lib.dom is missing size-adjust (supported since Chromium 92).
type FontFaceWidthDescriptors = FontFaceDescriptors & { sizeAdjust?: string | undefined }

function addFace(
  family: string,
  src: readonly string[],
  descriptors: FontFaceWidthDescriptors,
  loads: Promise<unknown>[],
): void {
  try {
    const face = new FontFace(family, faceSrc(src), descriptors)
    document.fonts.add(face)
    loads.push(face.load().catch(() => {}))
  } catch {
    /* invalid descriptor on this platform — keep the rest */
  }
}

function registerAlias(alias: CellFontAlias, loads: Promise<unknown>[]): void {
  let family = alias.family
  if (alias.scopeToCanvas) {
    family = `__cell-scope ${alias.family}`
    canvasScopedFamilies.set(alias.family.toLowerCase(), family)
  }
  addFace(family, alias.regular, { weight: '400', sizeAdjust: alias.sizeAdjust }, loads)
  if (alias.bold)
    addFace(family, alias.bold, { weight: '700', sizeAdjust: alias.boldSizeAdjust }, loads)
  const latin = alias.latin
  if (!latin) return
  const unicodeRange = latin.unicodeRange ?? LATIN_RANGE
  addFace(
    family,
    latin.regular,
    { weight: '400', sizeAdjust: latin.sizeAdjust, unicodeRange },
    loads,
  )
  if (latin.bold)
    addFace(
      family,
      latin.bold,
      { weight: '700', sizeAdjust: latin.boldSizeAdjust, unicodeRange },
      loads,
    )
}

/// Register + load alias faces before the first skeleton: canvas fillText
/// never triggers lazy @font-face loads, and Univer measures only once.
export function registerCellFontAliases(): Promise<unknown> {
  const loads: Promise<unknown>[] = []
  const gated: Promise<unknown>[] = []
  addFace(
    TEXT_DINGBATS_FAMILY,
    TEXT_DINGBATS_SOURCES,
    { weight: '400', unicodeRange: TEXT_DINGBATS_RANGE },
    loads,
  )
  for (const alias of CELL_FONT_ALIASES) {
    if (alias.skipIfLocal) {
      gated.push(
        genuineLocalExists(alias.skipIfLocal).then((genuine) => {
          const gatedLoads: Promise<unknown>[] = []
          if (genuine) {
            if (alias.whenGenuine)
              registerAlias({ family: alias.family, ...alias.whenGenuine }, gatedLoads)
          } else {
            substitutedFamilies.add(alias.family.toLowerCase())
            registerAlias(alias, gatedLoads)
          }
          return Promise.all(gatedLoads)
        }),
      )
    } else {
      registerAlias(alias, loads)
    }
  }
  return Promise.all([...loads, ...gated])
}
