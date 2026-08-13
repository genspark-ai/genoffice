/**
 * F1 line-box metrics engine
 *
 * Pure functions — input (paragraph text + run font/size + available width +
 * line-height rule + docGrid grid), output (line count + per-line heights + total block height).
 *
 * Design aligned with packages/pptx-render/src/metrics.ts (dual track: precise
 * opentype / heuristic fallback). F1 runs on heuristics first with the interface
 * ready, so we can switch seamlessly to OpentypeMetrics once font files load.
 *
 * Line-breaking rules:
 *   - Latin text breaks greedily by word (breakable at whitespace)
 *   - CJK breaks per character (Unicode CJK ranges)
 *   - Mixed text splits by Unicode segments (same logic as text-layout.ts, simplified port)
 *
 * Line-height semantics (Word's three modes):
 *   - auto    = multiple spacing; single = font natural line height (ascent+descent+lineGap)
 *   - atLeast = max(font natural line height, specified value (twips))
 *   - exact   = fixed value (twips), does not grow with the font
 *
 * docGrid:
 *   - with type='lines' or 'linesAndChars', each line height rounds up to a linePitch multiple
 *   - space-before/space-after keep their face value (Word does not grid-snap them)
 */

import type { DocGrid } from '@genoffice/docx-engine'

// ─── Font metrics interface (same interface as pptx-render/metrics.ts) ─────

export interface RunStyle {
  fontFamily: string
  fontSizePx: number
  bold: boolean
  italic: boolean
}

export interface FontMetrics {
  ascent: number
  descent: number
  lineHeight: number
}

export interface FontMetricsProvider {
  metrics(style: RunStyle): FontMetrics
  measure(text: string, style: RunStyle): number
}

// ─── Heuristic metrics (deterministic, unit-testable) ──────────────────────

/**
 * Estimate advance by character class (as a fraction of the font size).
 * Ported directly from pptx-render/metrics.ts HeuristicMetrics.
 */
const MONO_FONT_RE =
  /consolas|courier|menlo|monaco|cascadia|sf mono|jetbrains mono|source code|fira code|(liberation|roboto|ubuntu|dejavu sans|andale) mono|lucida console/i

/** Constant Latin advance for monospace families (chain-head name decides); null when proportional */
export function monospaceAdvanceEm(fontFamily: string): number | null {
  const head = fontFamily.split(',')[0]
  if (!MONO_FONT_RE.test(head)) return null
  return /consolas/i.test(head) ? 0.55 : 0.6
}

function charAdvanceEm(code: number): number {
  if (
    isHangul(code) ||
    (code >= 0x3000 && code <= 0x30ff) ||
    (code >= 0x3400 && code <= 0x9fff) ||
    (code >= 0xf900 && code <= 0xfaff) ||
    (code >= 0xff00 && code <= 0xff60) ||
    (code >= 0xffe0 && code <= 0xffe6)
  ) {
    return 1.0
  }
  if (code >= 0x1f000 || (code >= 0x2600 && code <= 0x27bf)) return 1.0
  if ("iIlj.,:;'!|".includes(String.fromCharCode(code))) return 0.28
  if (' ftr'.includes(String.fromCharCode(code))) return 0.32
  if ('mwMW'.includes(String.fromCharCode(code))) return 0.82
  return 0.52
}

export class HeuristicMetrics implements FontMetricsProvider {
  metrics(style: RunStyle): FontMetrics {
    const s = style.fontSizePx
    // Windows metric line heights ((usWinAscent+usWinDescent) / UPM) differ a lot by font:
    //   PMingLiU / MingLiU (Traditional Chinese): ~1.0em (compact)
    //   SimSun / NSimSun / SimHei families (Simplified Chinese): ~1.3em (expanded)
    //   Calibri / Arial / Times (Western): ~1.2em (standard)
    // ascent/descent keep the generic ratios (0.8/0.2 em); lineHeight uses the per-font factor
    return {
      ascent: s * 0.8,
      descent: s * 0.2,
      lineHeight: s * lineHeightFactor(style.fontFamily),
    }
  }

  measure(text: string, style: RunStyle): number {
    const monoEm = monospaceAdvanceEm(style.fontFamily)
    let em = 0
    for (const ch of text) {
      const cp = ch.codePointAt(0) ?? 0
      const base = charAdvanceEm(cp)
      // fullwidth (CJK/emoji) chars stay 1.0em even in monospace fonts
      em += monoEm !== null && base < 1 ? monoEm : base
    }
    const boldFactor = style.bold && monoEm === null ? 1.04 : 1
    return em * style.fontSizePx * boldFactor
  }
}

/** Korean font names (Windows/Noto/Source Han/Nanum faces + bundled subsets) */
const KO_FONT_RE =
  /malgun|맑은|batang|바탕|myeongjo|myungjo|명조|gungsuh|궁서|gulim|굴림|dotum|돋움|nanum|나눔|genoffice (sans|serif) kr|(noto|source han) (sans|serif)[^,]*\bk(r|orean)?\b/i

/**
 * Per-font single-spacing line-height factor. Re-baselined against Word for
 * Mac (batch probe 2026-08-13: docx per family/size converted by installed
 * Word, baseline pitch measured from the PDF). Latin faces matched the old
 * LO-era values exactly; the CJK entries moved a lot because Word renders
 * with its own bundled faces (SimSun/Batang/PMingLiU at 1.3029, Malgun
 * 1.7371, Meiryo 1.9429, YaHei 1.7143, Yu 1.44) instead of the macOS
 * substitutes LO measured. Unprobed names (SC/TC display classes, SimHei,
 * KaiTi) keep their LO values until probed.
 */
export function lineHeightFactor(fontFamily: string): number {
  const f = fontFamily.toLowerCase()
  // PMingLiU/MingLiU (Word probe: Office ships the real face at 1.3029)
  if (f.includes('pmingliu') || f.includes('mingliu') || f.includes('細明體')) {
    return 1.3029
  }
  // Korean faces (Word probe 2026-08-13: Malgun 1.7371, Batang class 1.3029)
  if (KO_FONT_RE.test(f)) return /malgun|맑은/.test(f) ? 1.7371 : 1.3029
  // Japanese faces: MS (P)Mincho/Gothic substitute into the Hiragino class
  // (1.7); Meiryo 1.775. Yu Mincho/Gothic: Word probe 2026-08-13 measured
  // exactly 1.44 at 10.5/12pt with Office's own yumin.ttf (the old 2.2667 was
  // an LO-baseline value and doubled every empty line under an 18pt docGrid)
  if (/游|yu (gothic|mincho)|yugoth|yumin/.test(f)) return 1.44
  if (/meiryo|メイリオ/.test(f)) return 1.9429
  if (/mincho|明朝|ゴシック|ms (ui )?p?gothic|hiragino|osaka|kozuka|小塚|biz ud/.test(f))
    return 1.3029
  // missing Noto/Source Han SC: Word substitutes SimSun at 1.3029 (probe
  // 2026-08-13; bare 'Noto Sans SC' presumed same substitution)
  if (/^noto sans sc$/.test(f)) return 1.3029
  if (/^(noto|source han) (sans|serif)( cjk)? ?(sc|cn|tc|tw|hk)\b/.test(f)) return 1.3029
  // Songti class (installed Songti SC / STSong)
  if (f.includes('simsun') || f.includes('nsimsun') || f.includes('宋体')) return 1.3029
  if (/(^|\s)(songti|stsong)\b/.test(f)) return 1.7
  if (/zhongsong|xiaobiaosong|中宋|小标宋/.test(f)) return 1.775
  // Compact macOS Heiti/Kaiti substitutes for SimHei and bare KaiTi names.
  if (f.includes('黑体') || f.includes('simhei')) return 1.0
  if ((f.includes('楷体') || f.includes('kaiti')) && !/gb2312|_gbk|gbk/.test(f)) return 1.0
  if (f.includes('microsoft yahei') || f.includes('microsoftyahei') || f.includes('雅黑'))
    return 1.7143
  // missing GB faces and other zh names substitute into the PingFang class
  if (
    f.includes('仿宋') ||
    f.includes('fangsong') ||
    f.includes('楷体') ||
    f.includes('kaiti') ||
    f.includes('宋') || // Song-family display faces and related variants
    f.includes('microsoft yahei') ||
    f.includes('microsoftyahei') ||
    f.includes('雅黑') ||
    f.includes('dengxian') ||
    f.includes('等线') ||
    f.includes('simfang') ||
    f.includes('simkai')
  ) {
    return 1.775
  }
  // Traditional Chinese sans/kai faces use the PingFang TC class.
  if (/jhenghei|正黑|標楷|biaukai|dfkai|kaiu/.test(f)) return 1.775
  // Arabic faces: Word for Mac substitutes missing naskh names with Times
  // New Roman (probe 2026-08-13)
  if (/naskh|kufi|arabic|amiri|scheherazade/.test(f)) return 1.1429
  // Western single-line factors follow each font's hhea metrics (LO probe):
  // a 4% surplus per line cascades into whole-paragraph pagination drift,
  // so the big Office faces get their real values.
  if (f.includes('times') || f.includes('liberation serif')) return 1.15
  if (f.includes('georgia')) return 1.1375
  if (f.includes('cambria') || f.includes('caladea')) return 1.17
  if (f.includes('helvetica')) return 1.0
  if (f === 'arial' || f.startsWith('arial ') || f.includes('liberation sans')) return 1.15
  if (f.includes('calibri') || f.includes('carlito')) return 1.22
  if (f.includes('tahoma')) return 1.2083
  if (f.includes('verdana')) return 1.2167
  // consolas first: the mono css chain carries 'Courier New' as fallback
  if (f.includes('consolas')) return 1.1667
  if (f.includes('courier')) return 1.1333
  if (f.includes('century') && !f.includes('gothic')) return 1.15
  if (f.includes('book antiqua')) return 1.1
  if (f.includes('segoe')) return 1.15
  if (/nyala|ebrima|abyssinica|ethiopic/.test(f)) return 1.0514
  // Tamil faces: Word renders missing Noto Sans Tamil / Latha with Latha
  // metrics (probe 2026-08-13)
  if (/tamil|latha|vijaya|inaimathi/.test(f)) return 1.6686
  // default (Lato, unknown Western)
  return 1.2
}

/**
 * Factor for run-declared Noto/Source Han regional variants missing on this
 * platform, for CJK text only. JP/SC/TC are Word-probed (2026-08-13): Word
 * for Mac substitutes SimSun at exactly 1.3029 single spacing (10/10.5pt,
 * grid and no-grid), explicit multiples scale linearly. KR reuses the
 * KO_FONT_RE value. Latin consumers (latinParaFactor, --doc-line-factor-latin)
 * must NOT use this: Cyrillic/Latin text under these names passes through
 * unicode-range aliases to proportional Latin glyphs where 1.2 matches Word
 * (sample 40).
 */
export function cjkDeclaredLineFactor(fontFamily: string): number | null {
  const m = /^(?:noto|source han) (?:sans|serif)(?: cjk)? ?(jp|kr|sc|cn|tc|tw|hk)\b/.exec(
    fontFamily.toLowerCase(),
  )
  if (!m) return null
  if (m[1] === 'kr') return lineHeightFactor(fontFamily)
  return 1.3029
}

/**
 * Word probe 2026-08-13: SimSun (the 1.3029 substitute for missing JP/SC/TC
 * faces) lacks ・ U+30FB and 〜 U+301C; Word falls back to Microsoft YaHei for
 * them and lifts the whole line to 1.7143 × size (13.68pt → 18.0pt @10.5pt).
 * KR (Batang class, also 1.3029) is unprobed and excluded.
 */
export const SIMSUN_GAP_CHAR_RE = /[・〜]/
export const SIMSUN_GAP_LINE_FACTOR = 1.7143

export function simsunGapLineFactor(fontFamily: string): number | null {
  const f = fontFamily.toLowerCase()
  if (KO_FONT_RE.test(f.normalize('NFKC'))) return null
  if (cjkDeclaredLineFactor(fontFamily) === 1.3029) return SIMSUN_GAP_LINE_FACTOR
  if (f.includes('simsun') || f.includes('nsimsun') || f.includes('宋体')) {
    return SIMSUN_GAP_LINE_FACTOR
  }
  return null
}

// ─── opentype.js precise metrics (interface ready; F1 falls back to heuristics) ─────

export interface OpentypeFontLike {
  unitsPerEm: number
  ascender: number
  descender: number
  getAdvanceWidth(text: string, fontSize: number): number
  charToGlyphIndex?(char: string): number
}

export class OpentypeMetrics implements FontMetricsProvider {
  constructor(
    private fontResolver: (style: RunStyle) => OpentypeFontLike | undefined,
    private fallback: FontMetricsProvider = new HeuristicMetrics(),
  ) {}

  metrics(style: RunStyle): FontMetrics {
    const font = this.fontResolver(style)
    if (!font) return this.fallback.metrics(style)
    const scale = style.fontSizePx / font.unitsPerEm
    const ascent = font.ascender * scale
    const descent = Math.abs(font.descender) * scale
    return { ascent, descent, lineHeight: ascent + descent }
  }

  measure(text: string, style: RunStyle): number {
    const font = this.fontResolver(style)
    if (!font) return this.fallback.measure(text, style)
    try {
      if (font.charToGlyphIndex) {
        for (const ch of text) {
          if (font.charToGlyphIndex(ch) === 0) return this.fallback.measure(text, style)
        }
      }
      return font.getAdvanceWidth(text, style.fontSizePx)
    } catch {
      return this.fallback.measure(text, style)
    }
  }
}

// ─── Line-height rule semantics ─────────────────────────────────────────────

const TWIPS_TO_PX = 96 / 1440

/** ε (fraction of a cell) so float noise just past a cell boundary stays in the lower cell */
const GRID_SNAP_EPS = 0.001

/**
 * Ceil a line height UP to whole grid cells. Single source for docGrid line
 * snapping: the pagination model calls it directly and cssGridLineExpr emits
 * the same formula (ε included) as a CSS round(up) — keep them in lockstep.
 */
export function snapLineToPitch(heightPx: number, pitchPx: number): number {
  return pitchPx > 0 ? Math.ceil(heightPx / pitchPx - GRID_SNAP_EPS) * pitchPx : heightPx
}

/**
 * Compute a single line's height (px) per Word's line-height rules.
 *
 * @param naturalLineH the font's natural line height (ascent+descent, px)
 * @param lineRule      'auto'|'atLeast'|'exact'
 * @param lineRawTwips  raw w:spacing w:line twips (auto = multiple of 240; atLeast/exact = absolute)
 * @param docGrid       the section docGrid (optional; when present, round to linePitch)
 */
export function computeLineHeight(
  naturalLineH: number,
  lineRule: 'auto' | 'atLeast' | 'exact' | undefined,
  lineRawTwips: number | undefined,
  docGrid: DocGrid | undefined,
): number {
  // typed line grid (LO probe 2026-08-10): single/auto lines snap UP to whole
  // linePitch cells BEFORE the auto multiple applies; exact/atLeast never snap
  const pitchPx =
    docGrid && (docGrid.type === 'lines' || docGrid.type === 'linesAndChars') && docGrid.linePitch
      ? docGrid.linePitch * TWIPS_TO_PX
      : 0
  const snapped = snapLineToPitch(naturalLineH, pitchPx)

  if (lineRule === 'exact' && lineRawTwips !== undefined) {
    // fixed line height: use the specified value regardless of font size
    return lineRawTwips * TWIPS_TO_PX
  }
  if (lineRule === 'atLeast' && lineRawTwips !== undefined) {
    // at least: max(font natural line height, specified value), unsnapped
    return Math.max(naturalLineH, lineRawTwips * TWIPS_TO_PX)
  }
  if (lineRule === 'auto' && lineRawTwips !== undefined) {
    // multiple: snapped base × (lineRawTwips / 240)
    return snapped * (lineRawTwips / 240)
  }
  // default (no lineRule): single spacing on the grid
  return snapped
}

/**
 * Font family → canvas CSS font-family fallback chain.
 * When a common Word font is missing locally, fall back to metric-compatible
 * open-source fonts (registered in fonts/fonts.css); identical widths keep line
 * breaks aligned with Word and the offline pagination model. CJK families fall
 * back to macOS equivalents (CJK width is always 1em, so this is mostly glyph appearance).
 */
/**
 * Whether a font family actually resolves on this machine. document.fonts.check
 * is useless in Chromium (true for any unknown system-ish family), so this
 * measures a mixed-script sample against both generic fallbacks: a family that
 * changes neither width doesn't exist.
 */
/**
 * Bundled web fonts (fonts.css). The canvas probe resolves them, but they are
 * last-resort subset faces, so substitution decisions treat them as missing.
 */
export const BUNDLED_FONTS = new Set([
  'Noto Sans CJK SC',
  'Noto Serif CJK SC',
  'GenOffice Sans KR',
  'GenOffice Serif KR',
  'GenOffice Fullwidth TC',
  'GenOffice Songti SC',
  'Carlito GO',
  'Caladea',
  'Liberation Serif',
  'Liberation Sans',
  'Liberation Mono',
])
const BUNDLED_LC = new Set([...BUNDLED_FONTS].map((f) => f.toLowerCase()))
export function isBundledFont(name: string): boolean {
  return BUNDLED_LC.has(name.toLowerCase())
}

const fontAvailableCache = new Map<string, boolean>()
export function isFontAvailable(font: string): boolean {
  if (typeof document === 'undefined') return false
  const cached = fontAvailableCache.get(font)
  if (cached !== undefined) return cached
  let available = false
  try {
    const canvas = document.createElement('canvas')
    const ctx = canvas.getContext('2d')
    if (ctx) {
      const sample = '한글あア中文abcWXYmm123'
      const quoted = `"${font.replace(/"/g, '')}"`
      const widthWith = (family: string) => {
        ctx.font = `32px ${family}`
        return ctx.measureText(sample).width
      }
      available =
        widthWith(`${quoted}, monospace`) !== widthWith('monospace') ||
        widthWith(`${quoted}, serif`) !== widthWith('serif')
    }
  } catch {
    /* headless/test environments treat every font as missing */
  }
  fontAvailableCache.set(font, available)
  return available
}

export function cssFontFamily(font: string): string {
  const f = font.toLowerCase()
  const chain = (...families: string[]) =>
    [...new Set(families)].map((x) => `'${x.replace(/'/g, '')}'`).join(',')
  // CJK fallback at chain end (GB2312 subset bundled in fonts.css): no tofu even without system Chinese fonts
  const CJK_SERIF = 'Noto Serif CJK SC'
  const CJK_SANS = 'Noto Sans CJK SC'
  // 'GenOffice Box Drawing' (fonts.css): U+2500 rules would otherwise fall to
  // the fullwidth CJK subset and overflow the column
  const BOX = 'GenOffice Box Drawing'
  if (f.includes('calibri')) return `${chain(font, 'Carlito GO', CJK_SANS)},sans-serif`
  // math faces would fall to the unknown-name sans fallback; STIX Two Math ships with macOS,
  // and on Windows the declared name resolves natively
  if ((f.includes('cambria') && f.includes('math')) || /stix.*math|latin modern math/.test(f))
    return `${chain(font, 'STIX Two Math', 'Caladea', 'Liberation Serif', BOX, CJK_SERIF)},serif`
  if (f.includes('cambria')) return `${chain(font, 'Caladea', BOX, CJK_SERIF)},serif`
  if (f.includes('times')) return `${chain(font, 'Liberation Serif', BOX, CJK_SERIF)},serif`
  if (f === 'arial' || f.startsWith('arial '))
    return `${chain(font, 'Liberation Sans', CJK_SANS)},sans-serif`
  if (MONO_FONT_RE.test(f))
    return `${chain(font, 'Menlo', 'Courier New', 'Liberation Mono', CJK_SANS)},monospace`
  // XiaoBiaoSong/ZhongSong (FZXiaoBiaoSong_GBK etc., gov-document title fonts) before the generic SimSun branch
  if (
    f.includes('小标宋') ||
    f.includes('中宋') ||
    f.includes('xiaobiaosong') ||
    f.includes('zhongsong')
  )
    return `${chain(font, 'STZhongsong', 'Songti SC', 'STSong', 'SimSun', CJK_SERIF)},serif`
  // 'GenOffice Songti SC' (fonts.css local() alias of Songti SC): macOS Chromium
  // refuses synthetic bold for 'Songti SC' by name at weight 600/700; the alias,
  // registered weight-normal only, lets Blink synthesize. Unresolvable elsewhere.
  if (f.includes('simsun') || f.includes('宋体') || f.includes('nsimsun')) {
    return `${chain(font, 'GenOffice Songti SC', 'STSong', 'SimSun', CJK_SERIF)},serif`
  }
  if (f.includes('simhei') || f.includes('黑体') || f.includes('细黑') || f.includes('xihei'))
    return `${chain(font, 'Heiti SC', 'STHeiti', 'SimHei', 'PingFang SC', CJK_SANS)},sans-serif`
  if (f.includes('yahei') || f.includes('雅黑'))
    return `${chain(font, 'Microsoft YaHei', 'PingFang SC', CJK_SANS)},sans-serif`
  if (f.includes('等线') || f.includes('dengxian'))
    return `${chain(font, 'DengXian', 'PingFang SC', 'Microsoft YaHei', CJK_SANS)},sans-serif`
  if (f.includes('fangsong') || f.includes('仿宋'))
    return `${chain(font, 'STFangsong', 'FangSong', CJK_SERIF)},serif`
  if (f.includes('kaiti') || f.includes('楷体'))
    return `${chain(font, 'STKaiti', 'Kaiti SC', 'KaiTi', CJK_SERIF)},serif`
  if (f.includes('隶书') || f.includes('lisu'))
    return `${chain(font, 'Baoli SC', 'LiSu', CJK_SERIF)},serif`
  // Japanese/Korean/Traditional Chinese: fall back within the same script (win/mac family names as mutual backups) so Han glyphs don't render with Simplified forms.
  // 'GenOffice *' entries are CJK-only local() aliases (fonts.css): the underlying
  // system faces draw Cyrillic/Greek fullwidth, so those scripts must pass through
  const JA_SANS = ['Yu Gothic', 'GenOffice Hiragino Sans', 'Meiryo', 'Noto Sans JP']
  const JA_SERIF = [
    'Yu Mincho',
    'GenOffice Hiragino Mincho',
    'GenOffice MS Mincho',
    'Noto Serif JP',
  ]
  const KO_SANS = ['Malgun Gothic', 'GenOffice Sans KR', 'Apple SD Gothic Neo', 'Noto Sans KR']
  const KO_SERIF = ['GenOffice Batang', 'GenOffice Serif KR', 'GenOffice Myungjo', 'Noto Serif KR']
  const TC_SANS = ['Microsoft JhengHei', 'PingFang TC', 'GenOffice Heiti TC', 'Noto Sans TC']
  // 'GenOffice Fullwidth TC' (fonts.css): fullwidth U+FF0D/FF0F/FF3C/FF3F/FF5E whose Songti TC glyphs look half-width
  const TC_SERIF = ['GenOffice MingLiU', 'GenOffice Fullwidth TC', 'Songti TC', 'Noto Serif TC']
  const SC_SANS = ['PingFang SC', 'Microsoft YaHei', CJK_SANS]
  const SC_SERIF = ['GenOffice Songti SC', 'STSong', 'SimSun', CJK_SERIF]
  const nfkc = font.normalize('NFKC')
  // Arabic: bundled Noto subsets stand in for missing fonts; Chromium's silent
  // fallback is a Geeza Pro-style UI face, larger and heavier than the naskh
  // serif Word substitutes. Unknown Arabic names default to the naskh chain.
  if (
    /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/.test(font) ||
    /naskh|kufi|arabic|urdu|geeza|amiri|scheherazade|lateef|harmattan|aldhabi|andalus|nastaliq|al bayan|baghdad|damascus|diwan|farisi|mishafi|nadeem|beirut/i.test(
      nfkc,
    )
  ) {
    const sans = /kufi|sans|dubai|segoe/i.test(nfkc)
    // Traditional/Simplified Arabic are compact naskh faces; the size-adjusted
    // alias (fonts.css) keeps advances near Word's, other Arabic names keep the
    // unscaled subset
    const compact = /\b(traditional|simplified) arabic\b/i.test(nfkc)
    // declared Noto Arabic names would resolve to the bundled unscaled subsets
    // (~13% wider than Word); the literal head must go so the 'W' alias wins
    const scaled = /^noto (naskh|sans) arabic$/i.test(nfkc.trim())
    const chainFor = sans
      ? [scaled ? 'Noto Sans Arabic W' : 'Noto Sans Arabic', 'Geeza Pro']
      : [
          compact ? 'Noto Naskh Arabic TA' : scaled ? 'Noto Naskh Arabic W' : 'Noto Naskh Arabic',
          'Geeza Pro',
          'Al Bayan',
        ]
    return `${chain(...(scaled ? [] : [font]), ...chainFor)},${sans ? 'sans-serif' : 'serif'}`
  }
  // Word substitutes a *missing* East Asian font with the locale's default face —
  // a serif (Mincho/Batang) — regardless of the requested font's classification.
  // Classic Windows faces (Malgun/Meiryo/Yu Gothic...) have solid macOS
  // equivalents in the chains and keep their classification.
  const missingLocally = () => isBundledFont(font) || !isFontAvailable(font)
  // Noto CJK / Source Han / Nanum regional variants route by suffix; the generic
  // fallback tail below would otherwise land them on the bundled Simplified-only subset
  const cjkVariant = /^(?:noto|source han) (sans|serif)(?: cjk)? ?(jp|kr|k|tc|tw|hk|sc|cn)\b/i.exec(
    nfkc,
  )
  if (cjkVariant) {
    const serif = /serif/i.test(cjkVariant[1]) || missingLocally()
    const region = cjkVariant[2].toLowerCase()
    const chainFor =
      region === 'jp'
        ? serif
          ? JA_SERIF
          : JA_SANS
        : region === 'kr' || region === 'k'
          ? serif
            ? KO_SERIF
            : KO_SANS
          : region === 'sc' || region === 'cn'
            ? serif
              ? SC_SERIF
              : SC_SANS
            : serif
              ? TC_SERIF
              : TC_SANS
    // a bundled face at the chain head would win over the substitution; it stays only as the tail safety net
    const head = isBundledFont(font) ? [] : [font]
    return `${chain(...head, ...chainFor)},${serif ? 'serif' : 'sans-serif'}`
  }
  if (
    /[぀-ヿ]|mincho|meiryo|hiragino|osaka|yugoth|yu (gothic|mincho)|ms (ui )?p?(gothic|mincho)|明朝|biz ud|kozuka|小塚/i.test(
      nfkc,
    )
  ) {
    const serif = /mincho|明朝/i.test(nfkc)
    return `${chain(font, ...(serif ? JA_SERIF : JA_SANS))},${serif ? 'serif' : 'sans-serif'}`
  }
  if (
    /[가-힣ᄀ-ᇿ㄰-㆏]|malgun|batang|gulim|dotum|gungsuh|myeongjo|myungjo|nanum|apple (sd )?gothic/i.test(
      nfkc,
    )
  ) {
    // vendor faces (Nanum...) missing locally follow Word's Batang-ward substitution;
    // Windows core faces (Malgun/Gulim/Dotum) map cleanly to the sans chain
    const knownCore = /malgun|맑은|gulim|굴림|dotum|돋움|apple (sd )?gothic/i.test(nfkc)
    const serif =
      /batang|바탕|myeongjo|myungjo|명조|gungsuh|궁서/i.test(nfkc) ||
      (!knownCore && missingLocally())
    return `${chain(font, ...(serif ? KO_SERIF : KO_SANS))},${serif ? 'serif' : 'sans-serif'}`
  }
  if (
    /jhenghei|p?mingliu|biaukai|dfkai|kaiu|正黑|細明|標楷|蘋方|-繁|繁體|pingfang (tc|hk)|(heiti|songti|kaiti) tc/i.test(
      nfkc,
    )
  ) {
    const serif = /mingliu|細明|標楷|biaukai|dfkai|kaiu|songti|kaiti|宋/i.test(nfkc)
    return `${chain(font, ...(serif ? TC_SERIF : TC_SANS))},${serif ? 'serif' : 'sans-serif'}`
  }
  // Ethiopic faces (Nyala/Ebrima/Abyssinica) are missing on macOS; the size-adjusted
  // Kefa alias (fonts.css) re-centers Chromium's ~1.35x-wide Kefa fallback. On
  // Windows the declared name resolves natively ahead of the alias.
  if (/nyala|ebrima|abyssinica|ethiopic/i.test(nfkc)) {
    return `${chain(font, 'GenOffice Ethiopic')},sans-serif`
  }
  // Tamil: macOS system faces stand in for Latha/Noto Sans Tamil; on Windows
  // the declared name resolves natively.
  // TODO: size-adjust alias to match Latha advances once sample 16 is re-shot
  if (/tamil|latha|vijaya|inaimathi/i.test(nfkc)) {
    return `${chain(font, 'InaiMathi', 'Tamil MN', 'Tamil Sangam MN')},sans-serif`
  }
  // unknown font family: guess serif-ness by name (Song/Ming/Serif → serif fallback)
  const serifLike = /宋|明|serif|song|ming/i.test(font)
  return `${chain(font, serifLike ? CJK_SERIF : CJK_SANS)},${serifLike ? 'serif' : 'sans-serif'}`
}

/**
 * Fallback chain for a dual-slot run (w:ascii ≠ w:eastAsia): Latin families first,
 * then the full East Asian chain. The Latin part drops its own CJK/generic fallbacks
 * so CJK glyphs fall through to the eastAsia font instead of a lookalike.
 */
export function cssDualFontFamily(ascii: string, eastAsia: string): string {
  if (ascii === eastAsia) return cssFontFamily(ascii)
  // Korean ascii face (e.g. theme latin = Malgun): its fallback chain covers hangul
  // and would swallow the eastAsia font, so keep only the literal family
  if (KO_FONT_RE.test(ascii.normalize('NFKC'))) {
    return `'${ascii.replace(/'/g, '')}',${cssFontFamily(eastAsia)}`
  }
  const latin = cssFontFamily(ascii)
    .split(',')
    .filter((f) => !/noto (sans|serif) cjk/i.test(f) && !/^(serif|sans-serif|monospace)$/.test(f))
  return `${latin.join(',')},${cssFontFamily(eastAsia)}`
}

/** Text contains complex-script characters (Arabic/Hebrew/Syriac/Thaana/NKo/Indic/Thai), i.e. the w:cs font slot applies */
export function textHasComplexScript(text: string): boolean {
  return /[\u0590-\u05FF\u0600-\u077F\u0780-\u07FF\u08A0-\u08FF\u0900-\u0DFF\u0E00-\u0E7F\uFB1D-\uFDFF\uFE70-\uFEFF]/.test(
    text,
  )
}

/**
 * Fallback chain for a run whose text hits the w:cs slot: the cs chain leads
 * (minus its generic tail) so complex-script glyphs use it, then the run's
 * Latin/East Asian chain for everything else.
 */
export function cssCsFontFamily(cs: string, ascii?: string, eastAsia?: string): string {
  const base =
    ascii && eastAsia && ascii !== eastAsia
      ? cssDualFontFamily(ascii, eastAsia)
      : ascii || eastAsia
        ? cssFontFamily((ascii || eastAsia)!)
        : ''
  if (!base) return cssFontFamily(cs)
  const head = cssFontFamily(cs)
    .split(',')
    .filter((f) => !/^(serif|sans-serif|monospace)$/.test(f))
  const baseFams = base.split(',')
  const generic = /^(serif|sans-serif|monospace)$/
  // Bundled Noto Arabic subsets have no Latin letters/parens; splice the run's
  // Latin chain in right after them, before Geeza Pro whose Latin punctuation
  // sits on the Arabic baseline. Arabic glyphs resolve in the subset first, so
  // shaping is unaffected.
  const notoIdx = head.findIndex((f) => /noto (naskh|sans) arabic/i.test(f))
  const merged =
    notoIdx >= 0
      ? [
          ...head.slice(0, notoIdx + 1),
          ...baseFams.filter((f) => !generic.test(f)),
          ...head.slice(notoIdx + 1),
          ...baseFams.filter((f) => generic.test(f)),
        ]
      : [...head, ...baseFams]
  return [...new Set(merged)].join(',')
}

/** Text contains CJK characters (decides the line-height factor: CJK lines measure ~1.3em per Chinese font metrics) */
export function textHasCjk(text: string): boolean {
  for (const ch of text) {
    if (isCjk(ch.codePointAt(0) ?? 0)) return true
  }
  return false
}

/** Text contains hangul (Korean paragraphs take the Korean line factor, not the 1.3 Chinese one) */
export function textHasHangul(text: string): boolean {
  for (const ch of text) {
    if (isHangul(ch.codePointAt(0) ?? 0)) return true
  }
  return false
}

/** Per-paragraph --doc-line-factor value by script (approximates Word's max-of-inline-fonts line height) */
export function paraLineFactorCss(text: string): string {
  if (textHasHangul(text)) return 'var(--doc-line-factor-kr,1.3029)'
  if (textHasCjk(text)) return 'var(--doc-line-factor-cjk,1.7)'
  return 'var(--doc-line-factor-latin,1.2)'
}

/** --doc-line-factor-kr source: the document's East Asian face when Korean, else the Batang-class default */
export function krLineFactor(fontFamily: string | undefined): number {
  return fontFamily && isKoreanFontName(fontFamily) ? lineHeightFactor(fontFamily) : 1.3029
}

export function isKoreanFontName(fontFamily: string): boolean {
  return KO_FONT_RE.test(fontFamily.normalize('NFKC'))
}

/**
 * Word line spacing → editing-canvas CSS line-height value (shared by extensions/docStyleCss).
 *
 * Line rules: auto = multiple × font natural line height; atLeast = max(natural, N); exact = N.
 * Auto/multiple emits a UNITLESS number (var(--doc-line-factor) × multiple):
 * numbers inherit by value and rescale with each run's own font size, matching
 * Word's tallest-run-per-line rule (a length computed at the paragraph would
 * inherit into larger runs and collapse their lines).
 *
 * docGrid (LO probe, 2026-08-10): in sections with a typed line grid
 * (w:docGrid lines/linesAndChars) single/auto-multiple lines snap UP to a
 * whole number of linePitch cells and the multiple applies AFTER snapping;
 * exact and atLeast lines never snap; w:snapToGrid=0 opts a paragraph out.
 * round() needs a length, so grid snapping cannot be unitless: docStyleCss
 * declares --doc-line-grid (cssGridLineExpr) on every element of typed-grid
 * documents, and auto/multiple line heights resolve it ahead of the unitless
 * fallback. Per-element declaration keeps per-paragraph --doc-line-factor /
 * --doc-grid-pitch overrides live (a value inherited from .doc-page would
 * freeze the vars substituted there).
 */
const GRID_PITCH = 'var(--doc-grid-pitch,0.0001px)'

/** grid-snapped single-line height; docStyleCss assigns it to --doc-line-grid in typed-grid docs.
 *  CSS mirror of snapLineToPitch (same ε); 1em is the paragraph's own font size —
 *  Word snaps by the tallest run per line, one line-height per paragraph is our simplification. */
export function cssGridLineExpr(): string {
  return `round(up, calc(var(--doc-line-factor,1.2) * 1em - ${GRID_PITCH} * ${GRID_SNAP_EPS}), ${GRID_PITCH})`
}

/** SimSun-gap lifted line (・/〜 runs): multiple × grid-snapped 1.7143em, mirroring the
 *  pagination model (snap the boosted natural height, then apply the multiple) */
export function cssSimsunGapLineExpr(multiple: number): string {
  const single = `round(up, calc(${SIMSUN_GAP_LINE_FACTOR} * 1em - ${GRID_PITCH} * ${GRID_SNAP_EPS}), ${GRID_PITCH})`
  return multiple === 1 ? single : `calc(${single} * ${multiple})`
}

/** single-spacing line height: snapped length in typed-grid docs, unitless factor otherwise */
export function cssGridLineBase(): string {
  return 'var(--doc-line-grid,var(--doc-line-factor,1.2))'
}

/**
 * Paragraph space before/after (pt) as a CSS margin value. Word keeps
 * paragraph spacing at face value even under a typed grid (probe 2026-08-13:
 * before 4.5/10/18/24pt all render as written on an 18pt grid); the earlier
 * round-down-to-grid-cells behavior was an LO artifact.
 */
export function cssGridSpacingPt(pt: number): string {
  return `${pt.toFixed(1)}pt`
}

export function cssLineHeight(
  lineRule: 'auto' | 'atLeast' | 'exact' | undefined,
  lineRawTwips: number | undefined,
  lineSpacing: number | undefined,
): string | null {
  const FACTOR = 'var(--doc-line-factor,1.2)'
  if (lineRule === 'exact' && lineRawTwips) return `${(lineRawTwips / 20).toFixed(1)}pt`
  if (lineRule === 'atLeast' && lineRawTwips != null) {
    // atLeast never snaps to the grid; line="0" atLeast degrades to plain natural height
    if (lineRawTwips === 0) return `calc(${FACTOR} * 1em)`
    return `max(${(lineRawTwips / 20).toFixed(1)}pt, calc(${FACTOR} * 1em))`
  }
  const m = lineSpacing ?? (lineRule === 'auto' && lineRawTwips ? lineRawTwips / 240 : undefined)
  if (m) return m === 1 ? cssGridLineBase() : `calc(${cssGridLineBase()} * ${m})`
  return null
}

/**
 * Space-before/space-after in px. Word keeps paragraph spacing at face value
 * even under a typed grid (probe 2026-08-13), so no grid quantization.
 */
export function snapSpacingToGrid(spacingTwips: number, _docGrid: DocGrid | undefined): number {
  return spacingTwips * TWIPS_TO_PX
}

// ─── CJK detection ───────────────────────────────────────────────────────────

function isCjk(cp: number): boolean {
  return (
    isHangul(cp) ||
    (cp >= 0x3000 && cp <= 0x30ff) ||
    (cp >= 0x3400 && cp <= 0x9fff) ||
    (cp >= 0xf900 && cp <= 0xfaff) ||
    (cp >= 0xff00 && cp <= 0xffef)
  )
}

function isHangul(cp: number): boolean {
  return (
    (cp >= 0xac00 && cp <= 0xd7a3) ||
    (cp >= 0x1100 && cp <= 0x11ff) ||
    (cp >= 0x3130 && cp <= 0x318f) ||
    (cp >= 0xa960 && cp <= 0xa97f) ||
    (cp >= 0xd7b0 && cp <= 0xd7ff)
  )
}

// ─── CJK-Latin autospace pads ────────────────────────────────────────────────
// Word's autoSpaceDE/DN gap measures ~1/4em while Chromium's text-autospace is
// fixed at 1/8em; the renderer inserts a zero-width .doc-autospace-pad span
// whose margin supplies the other 1/8em. Pads only go between a directly
// adjacent CJK letter and a Latin letter/digit — never next to spaces or
// punctuation — matching where Chromium applies its native gap.

/** Han/kana/hangul letters; CJK punctuation and full/halfwidth forms get no gap */
function isCjkAutospaceSide(cp: number): boolean {
  if (cp >= 0x3000 && cp <= 0x303f) return false
  if (cp === 0x30fb) return false
  if (cp >= 0xff00 && cp <= 0xffef) return false
  return isCjk(cp)
}

function isLatinAlnum(cp: number): boolean {
  if ((cp >= 0x30 && cp <= 0x39) || (cp >= 0x41 && cp <= 0x5a) || (cp >= 0x61 && cp <= 0x7a)) {
    return true
  }
  // Latin-1 Supplement / Extended letters, minus multiply/divide signs
  return cp >= 0xc0 && cp <= 0x24f && cp !== 0xd7 && cp !== 0xf7
}

export function needsAutospacePad(prevCp: number, nextCp: number): boolean {
  return (
    (isCjkAutospaceSide(prevCp) && isLatinAlnum(nextCp)) ||
    (isLatinAlnum(prevCp) && isCjkAutospaceSide(nextCp))
  )
}

function lastCodePoint(text: string): number {
  const tail = text.charCodeAt(text.length - 1)
  if (tail >= 0xdc00 && tail <= 0xdfff && text.length > 1) {
    return text.codePointAt(text.length - 2)!
  }
  return tail
}

/** pad between two adjacent stretches of text (last char of prev vs first of next) */
export function autospacePadBetween(prev: string, next: string): boolean {
  if (!prev || !next) return false
  return needsAutospacePad(lastCodePoint(prev), next.codePointAt(0)!)
}

/** UTF-16 offsets inside text where a pad belongs (between offset-1 and offset) */
export function autospaceBoundaries(text: string): number[] {
  const out: number[] = []
  let prev = -1
  for (let i = 0; i < text.length;) {
    const cp = text.codePointAt(i)!
    if (prev >= 0 && needsAutospacePad(prev, cp)) out.push(i)
    prev = cp
    i += cp > 0xffff ? 2 : 1
  }
  return out
}

// ─── Line-break simulation ───────────────────────────────────────────────────

/**
 * CJK character line-height factor (by font family).
 * Known CJK families take their LO-probe factor (lineHeightFactor); CJK text
 * inside Latin-font runs falls through LO's font fallback (~1.3 on this
 * platform), so the legacy 1.3 stays as the default.
 */
/** Font name plausibly covering CJK glyphs (drives whether a declared face may
 *  set the CJK line factor; Latin names cascade to the document's EA default). */
export function isCjkFontName(fontFamily: string): boolean {
  const f = fontFamily.toLowerCase()
  return CJK_FONT_NAME_RE.test(f) || KO_FONT_RE.test(f.normalize('NFKC'))
}

const CJK_FONT_NAME_RE =
  /宋|黑|楷|仿|明|雅黑|等线|simsun|simhei|simkai|simfang|kaiti|fangsong|songti|stsong|yahei|dengxian|mingliu|jhenghei|biaukai|dfkai|kaiu|mincho|ms (ui )?p?gothic|yu gothic|ゴシック|meiryo|メイリオ|hiragino|osaka|kozuka|游|(noto|source han) (sans|serif)( cjk)? ?(sc|cn|jp|tc|kr)\b/i

function cjkLineHFactor(fontFamily: string): number {
  const f = fontFamily.toLowerCase()
  if (f.includes('pmingliu') || f.includes('mingliu')) return 1.0
  const declared = cjkDeclaredLineFactor(fontFamily)
  if (declared !== null) return declared
  if (KO_FONT_RE.test(f)) return lineHeightFactor(fontFamily)
  if (CJK_FONT_NAME_RE.test(f)) return lineHeightFactor(fontFamily)
  return 1.3
}

/**
 * Paragraph line-break simulation — computes line count (heights only, no glyph coordinates).
 *
 * Logic aligned with text-layout.ts layoutParagraph, but outputs only the list of line boxes (heights).
 * Each item in the runs array is a stretch of text with the same style (same as Block.runs).
 *
 * cjkFactor semantics:
 *   NaN (default) = decided dynamically per font family via cjkLineHFactor(style.fontFamily)
 *   0..N = fixed factor (table cells pass 1.0 = no extra boost)
 */
const CJK_LINE_HEIGHT_FACTOR = NaN

export function simulateLines(
  runs: Array<{
    text: string
    fontFamily?: string
    sizeHalfPoints?: number
    bold?: boolean
    italic?: boolean
  }>,
  availWidthPx: number,
  metrics: FontMetricsProvider,
  defaultFontSize: number,
  defaultFontFamily: string,
  cjkFactor = CJK_LINE_HEIGHT_FACTOR,
): Array<{ naturalLineH: number; text: string }> {
  if (availWidthPx <= 0)
    return [{ naturalLineH: defaultFontSize * 1.2, text: runs.map((r) => r.text).join('') }]

  // Word breaks Korean at spaces (keep-all), including Han inside Korean
  // paragraphs; the renderer applies the matching word-break:keep-all CSS
  const keepAll = runs.some((r) => textHasHangul(r.text))

  // resulting line list (records each line's natural height and text)
  const lines: Array<{ naturalLineH: number; text: string }> = []
  let curLineW = 0
  let curLineH = 0 // max natural line height of the current line
  let curText = ''

  const pushLine = (h: number) => {
    lines.push({ naturalLineH: h, text: curText })
    curLineW = 0
    curLineH = 0
    curText = ''
  }

  const getStyle = (run: (typeof runs)[0]): RunStyle => ({
    fontFamily: run.fontFamily ?? defaultFontFamily,
    fontSizePx: (run.sizeHalfPoints ? run.sizeHalfPoints / 2 : defaultFontSize) * (96 / 72),
    bold: !!run.bold,
    italic: !!run.italic,
  })

  for (const run of runs) {
    if (!run.text) continue
    const style = getStyle(run)
    const m = metrics.metrics(style)
    // table-cell mode (cjkFactor=1.0): line-height cap = fontSizePx × 1.0,
    // matching the CJK chars' cjkH so Latin chars' font factor (1.2) doesn't raise the line.
    const lineH = isNaN(cjkFactor)
      ? m.lineHeight
      : Math.min(m.lineHeight, style.fontSizePx * Math.max(cjkFactor, 1.0))

    let buf = ''
    let bufCjkH = 0 // CJK line-height floor of buffered chars (keep-all mode)
    const flushWord = () => {
      if (!buf) return
      const wordH = Math.max(lineH, bufCjkH)
      bufCjkH = 0
      const w = metrics.measure(buf, style)
      if (curLineW + w > availWidthPx && curLineW > 0) {
        pushLine(Math.max(curLineH, lineH))
        curLineH = lineH
      }
      if (curLineW === 0 && w > availWidthPx) {
        // hard-break an overlong word
        let fragment = ''
        for (const ch of buf) {
          const cw = metrics.measure(fragment + ch, style)
          if (fragment && cw > availWidthPx) {
            curText = fragment
            pushLine(wordH)
            curLineH = wordH
            fragment = ch
          } else {
            fragment += ch
          }
        }
        if (fragment) {
          curLineW += metrics.measure(fragment, style)
          curLineH = Math.max(curLineH, wordH)
          curText = fragment
        }
      } else {
        curLineW += w
        curLineH = Math.max(curLineH, wordH)
        curText += buf
      }
      buf = ''
    }

    for (const ch of run.text) {
      const cp = ch.codePointAt(0) ?? 0
      if (ch === '\n') {
        flushWord()
        pushLine(Math.max(curLineH, lineH))
        curLineH = lineH
        continue
      }
      if (ch === ' ' || ch === '\t') {
        flushWord()
        const spW = metrics.measure(ch, style)
        if (curLineW + spW > availWidthPx && curLineW > 0) {
          pushLine(Math.max(curLineH, lineH))
          curLineH = lineH
          // swallow the space at line start
        } else {
          curLineW += spW
          curLineH = Math.max(curLineH, lineH)
          curText += ch
        }
        continue
      }
      if (isCjk(cp)) {
        // CJK char line height: NaN = picked dynamically per font family (body path), a number = fixed factor (table cells)
        let runCjkFactor = isNaN(cjkFactor) ? cjkLineHFactor(style.fontFamily) : cjkFactor
        // ・/〜 under SimSun substitution lift the line to 1.7143 (probe 2026-08-13)
        if (isNaN(cjkFactor) && SIMSUN_GAP_CHAR_RE.test(ch)) {
          runCjkFactor = Math.max(runCjkFactor, simsunGapLineFactor(style.fontFamily) ?? 0)
        }
        const cjkH = style.fontSizePx * runCjkFactor
        if (keepAll) {
          buf += ch
          bufCjkH = Math.max(bufCjkH, cjkH)
          continue
        }
        flushWord()
        const cw = metrics.measure(ch, style)
        if (curLineW + cw > availWidthPx && curLineW > 0) {
          pushLine(Math.max(curLineH, lineH))
          curLineH = lineH
        }
        curLineW += cw
        curText += ch
        curLineH = Math.max(curLineH, cjkH)
        continue
      }
      buf += ch
    }
    flushWord()
  }

  // wrap up (the last line)
  if (curLineH > 0 || lines.length === 0) {
    // last-line minimum:
    //   body path (cjkFactor=NaN): use the font-level line-height factor (lineHeightFactor)
    //   table-cell path (cjkFactor=1.0): match CJK chars, using cjkFactor as the factor
    const lastLineMin = isNaN(cjkFactor)
      ? defaultFontSize * (96 / 72) * lineHeightFactor(defaultFontFamily)
      : defaultFontSize * (96 / 72) * Math.max(cjkFactor, 1.0)
    pushLine(Math.max(curLineH, lastLineMin))
  }

  return lines
}

/**
 * Widest unbreakable token width (px) across runs: whitespace splits words;
 * CJK breaks per character unless Hangul is present (keep-all, same rule as
 * simulateLines). Min-content input for table autofit; scanning is capped so
 * huge cells stay cheap.
 */
export function maxWordWidthPx(
  runs: Array<{
    text: string
    fontFamily?: string
    sizeHalfPoints?: number
    bold?: boolean
    italic?: boolean
  }>,
  metrics: FontMetricsProvider = defaultMetrics,
  maxWords = 200,
): number {
  const keepAll = runs.some((r) => textHasHangul(r.text))
  let max = 0
  let cur = 0 // width of the current token's completed segments (may span runs)
  let words = 0
  let scanned = 0
  const endWord = () => {
    if (cur > max) max = cur
    if (cur > 0) words++
    cur = 0
  }
  for (const run of runs) {
    if (!run.text) continue
    const style: RunStyle = {
      fontFamily: run.fontFamily ?? DEFAULT_FONT_FAMILY,
      fontSizePx: (run.sizeHalfPoints ? run.sizeHalfPoints / 2 : DEFAULT_FONT_SIZE_PT) * (96 / 72),
      bold: !!run.bold,
      italic: !!run.italic,
    }
    let buf = ''
    const flushSeg = () => {
      if (!buf) return
      cur += metrics.measure(buf, style)
      buf = ''
    }
    for (const ch of run.text) {
      if (words >= maxWords || scanned >= maxWords * 60) {
        flushSeg()
        endWord()
        return max
      }
      scanned++
      if (ch === ' ' || ch === '\t' || ch === '\n') {
        flushSeg()
        endWord()
        continue
      }
      if (!keepAll && isCjk(ch.codePointAt(0) ?? 0)) {
        flushSeg()
        endWord()
        cur = metrics.measure(ch, style)
        endWord()
        continue
      }
      buf += ch
    }
    flushSeg()
  }
  endWord()
  return max
}

// ─── Main entry point ────────────────────────────────────────────────────────

export interface LineMetricsInput {
  runs: Array<{
    text: string
    fontFamily?: string
    sizeHalfPoints?: number
    bold?: boolean
    italic?: boolean
  }>
  availWidthPx: number
  /** paragraph line-height rule */
  lineRule?: 'auto' | 'atLeast' | 'exact'
  /** raw w:spacing w:line twips */
  lineRawTwips?: number
  /** space before, twips */
  spaceBefore?: number
  /** space after, twips */
  spaceAfter?: number
  /** section docGrid (affects line-height rounding) */
  docGrid?: DocGrid
  /** document default font size (pt, from docDefaults.sizeHalfPoints/2 or fallback 12) */
  defaultFontSizePt?: number
  /** document default font */
  defaultFontFamily?: string
  /** metrics implementation (default HeuristicMetrics) */
  metrics?: FontMetricsProvider
  /** whether the paragraph is empty (a textless paragraph occupies one line height) */
  isEmpty?: boolean
  /**
   * Table-cell mode: disables the CJK line-height boost (1.5× → 1.2×).
   * Cell line height is controlled by the line-height rule + docGrid, without extra CJK expansion.
   */
  tableCellMode?: boolean
  /** CJK line-height factor override (e.g. tuned against the baseline layout engine); defaults to tableCellMode/font-family behavior */
  cjkFactor?: number
}

export interface LineMetricsResult {
  /** line count */
  lineCount: number
  /** per-line heights (px, with line-height rules and docGrid rounding applied) */
  lineHeights: number[]
  /** per-line text (aligned with lineHeights; used to locate the page-leading line for in-block page splits) */
  lineTexts: string[]
  /** space before (px) */
  spaceBeforePx: number
  /** space after (px) */
  spaceAfterPx: number
  /** total block height = sum(lineHeights) + spaceBeforePx + spaceAfterPx */
  totalHeight: number
}

/**
 * Line box (for line-level page splitting).
 * offsetInBlock: offset of the line's top relative to the paragraph block's top, including spaceBefore (px).
 * height: the line's height (px).
 */
export interface LineBox {
  /** top offset of the line within the block (px, relative to block top) */
  offsetInBlock: number
  /** line height (px) */
  height: number
}

/**
 * Extended computeLineMetrics: additionally returns each line's LineBox (with offsets).
 * Used for line-level split decisions — the pagination algorithm can break a paragraph at any line boundary.
 */
export interface LineMetricsResultEx extends LineMetricsResult {
  /** per-line boxes (with in-block offsets), consumed by line-level page splitting */
  lineBoxes: LineBox[]
}

const DEFAULT_FONT_FAMILY = 'Calibri'
const DEFAULT_FONT_SIZE_PT = 12

/** global default HeuristicMetrics instance (avoids rebuilding per call) */
const defaultMetrics = new HeuristicMetrics()

/**
 * Compute a paragraph's precise line-box metrics (block height from Word's perspective).
 *
 * Used for pagination computation (replacing getBoundingClientRect); does not affect editor rendering.
 */
export function computeLineMetrics(input: LineMetricsInput): LineMetricsResultEx {
  const {
    runs,
    availWidthPx,
    lineRule,
    lineRawTwips,
    spaceBefore = 0,
    spaceAfter = 0,
    docGrid,
    defaultFontSizePt = DEFAULT_FONT_SIZE_PT,
    defaultFontFamily = DEFAULT_FONT_FAMILY,
    metrics = defaultMetrics,
    isEmpty = false,
    tableCellMode = false,
  } = input

  const defaultFontSizePx = defaultFontSizePt * (96 / 72)
  // table cells disable the extra CJK line-height boost (cjkFactor=1.0, relying on HeuristicMetrics' font-level line height)
  // body paragraphs use the font-level cjkLineHFactor (PMingLiU→1.0, others→1.3), triggered via NaN
  const cjkFactor = input.cjkFactor ?? (tableCellMode ? 1.0 : CJK_LINE_HEIGHT_FACTOR)

  // natural line height of an empty paragraph: use the font-level factor (no hard-coded 1.2)
  const emptyNaturalH = defaultFontSizePx * lineHeightFactor(defaultFontFamily)

  // simulate line breaking
  const lines =
    isEmpty || runs.length === 0
      ? [{ naturalLineH: emptyNaturalH, text: '' }]
      : simulateLines(runs, availWidthPx, metrics, defaultFontSizePt, defaultFontFamily, cjkFactor)

  // apply the line-height rule per line
  const lineHeights = lines.map((ln) =>
    computeLineHeight(ln.naturalLineH, lineRule, lineRawTwips, docGrid),
  )

  // space before/after
  // table-cell mode: no docGrid alignment for paragraph spacing (only line heights snap to the grid)
  const spacingDocGrid = tableCellMode ? undefined : docGrid
  const spaceBeforePx = spaceBefore > 0 ? snapSpacingToGrid(spaceBefore, spacingDocGrid) : 0
  const spaceAfterPx = spaceAfter > 0 ? snapSpacingToGrid(spaceAfter, spacingDocGrid) : 0

  const totalHeight = lineHeights.reduce((s, h) => s + h, 0) + spaceBeforePx + spaceAfterPx

  // build each line's LineBox (with in-block offsets)
  // spaceBefore counts before the first line
  const lineBoxes: LineBox[] = []
  let offset = spaceBeforePx
  for (const h of lineHeights) {
    lineBoxes.push({ offsetInBlock: offset, height: h })
    offset += h
  }

  return {
    lineCount: lineHeights.length,
    lineHeights,
    lineTexts: lines.map((ln) => ln.text),
    lineBoxes,
    spaceBeforePx,
    spaceAfterPx,
    totalHeight,
  }
}

/** reserved height for the footnote separator line (px) */
export const FOOTNOTE_SEPARATOR_H = 16

/**
 * Header/footer part height estimate (px): per-paragraph line-box model. Capacity
 * input for body push-down (body top = max(marginTop, headerDist + header height)).
 */
export function estimateHfHeight(
  part:
    | {
        text: string
        paras?: Array<{
          runs: Array<{
            text: string
            font?: string
            sizeHalfPoints?: number
            bold?: boolean
            italic?: boolean
          }>
          lineRule?: 'auto' | 'atLeast' | 'exact'
          lineRawTwips?: number
          lineSpacing?: number
          spaceBefore?: number
          spaceAfter?: number
        }>
      }
    | null
    | undefined,
  contentWidthPx: number,
  /** images in the part (logos): non-floating ones count as one line of height (same as the display layer) */
  images?: Array<{ heightPx?: number; floating?: boolean }> | null,
): number {
  const inlineImages = (images ?? []).filter((im) => !im.floating && im.heightPx)
  const imagesHeight =
    inlineImages.length > 0 ? Math.max(...inlineImages.map((im) => im.heightPx!)) + 2 : 0
  if (!part) return imagesHeight
  type HfPara = NonNullable<typeof part.paras>[number]
  const paras: HfPara[] = part.paras?.length
    ? part.paras
    : part.text.trim()
      ? part.text.split('\n').map((t) => ({ runs: [{ text: t }] }))
      : []
  let height = 0
  for (const p of paras) {
    const rich = p as NonNullable<typeof part.paras>[number]
    height += computeLineMetrics({
      runs: p.runs.map((r) => ({
        text: r.text,
        ...(r.font ? { fontFamily: r.font } : {}),
        ...(r.sizeHalfPoints ? { sizeHalfPoints: r.sizeHalfPoints } : {}),
        ...(r.bold ? { bold: true } : {}),
        ...(r.italic ? { italic: true } : {}),
      })),
      availWidthPx: contentWidthPx,
      ...(rich.lineRule ? { lineRule: rich.lineRule } : {}),
      ...(rich.lineRawTwips
        ? { lineRawTwips: rich.lineRawTwips }
        : rich.lineSpacing
          ? { lineRule: 'auto' as const, lineRawTwips: rich.lineSpacing * 240 }
          : {}),
      ...(rich.spaceBefore ? { spaceBefore: rich.spaceBefore } : {}),
      ...(rich.spaceAfter ? { spaceAfter: rich.spaceAfter } : {}),
      defaultFontSizePt: 10.5,
      isEmpty: p.runs.every((r) => !r.text.trim()),
    }).totalHeight
  }
  return height + imagesHeight
}

/**
 * Footnote entry height estimate (10pt small text, content width narrowed 40px to approximate Word's footnote layout).
 * The referencing page reserves bottom space accordingly (same model as the parity runner).
 * The 40px narrowing is wider than the renderer's sup-number prefix, so wrap
 * mismatches err toward over-reserving (never clipping).
 */
export function estimateFootnoteHeight(
  footnoteText: string,
  contentWidthPx: number,
  docGrid: DocGrid | undefined,
  metrics?: FontMetricsProvider,
): number {
  return computeLineMetrics({
    runs: [{ text: footnoteText }],
    availWidthPx: contentWidthPx - 40,
    docGrid,
    defaultFontSizePt: 10,
    ...(metrics ? { metrics } : {}),
  }).totalHeight
}

/**
 * Line height (px) of the footnote estimate model above (10pt default font,
 * docGrid-snapped). The gap-notes renderer sets this as its CSS line-height so
 * estimate and rendering share one line-box truth.
 */
export function footnoteLineHeightPx(docGrid: DocGrid | undefined): number {
  const naturalH = 10 * (96 / 72) * lineHeightFactor(DEFAULT_FONT_FAMILY)
  return computeLineHeight(naturalH, undefined, undefined, docGrid)
}
