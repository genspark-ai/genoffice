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
 *   - space-before/space-after also align to the grid when a grid exists
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
function charAdvanceEm(code: number): number {
  if (
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
    let em = 0
    for (const ch of text) {
      const cp = ch.codePointAt(0) ?? 0
      em += charAdvanceEm(cp)
    }
    const boldFactor = style.bold ? 1.04 : 1
    return em * style.fontSizePx * boldFactor
  }
}

/**
 * Per-font line-height factor (based on measured Windows font metrics).
 * PMingLiU/MingLiU: ~1.0em; SimSun and other Simplified fonts: ~1.3em; Western: 1.2em.
 */
export function lineHeightFactor(fontFamily: string): number {
  const f = fontFamily.toLowerCase()
  // compact Traditional Chinese fonts (PMingLiU, MingLiU, etc.)
  if (f.includes('pmingliu') || f.includes('mingliu') || f.includes('細明體')) {
    return 1.0
  }
  // expanded Simplified Chinese fonts (SimSun/NSimSun/SimHei/FangSong/KaiTi/Microsoft YaHei)
  if (
    f.includes('simsun') ||
    f.includes('nsimsun') ||
    f.includes('宋体') ||
    f.includes('黑体') ||
    f.includes('fangsong') ||
    f.includes('kaiti') ||
    f.includes('microsoft yahei') ||
    f.includes('microsoftyahei') ||
    f.includes('simhei') ||
    f.includes('simfang') ||
    f.includes('simkai')
  ) {
    return 1.3
  }
  // Western and default (Calibri, Arial, Times, Liberation, etc.)
  return 1.2
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
  let lineH: number

  if (lineRule === 'exact' && lineRawTwips !== undefined) {
    // fixed line height: use the specified value regardless of font size
    lineH = lineRawTwips * TWIPS_TO_PX
  } else if (lineRule === 'atLeast' && lineRawTwips !== undefined) {
    // at least: max(font natural line height, specified value)
    lineH = Math.max(naturalLineH, lineRawTwips * TWIPS_TO_PX)
  } else if (lineRule === 'auto' && lineRawTwips !== undefined) {
    // multiple: naturalLineH * (lineRawTwips / 240)
    lineH = naturalLineH * (lineRawTwips / 240)
  } else {
    // default (no lineRule): single spacing
    lineH = naturalLineH
  }

  // docGrid: real Word for Mac baseline (2026-07-27, line-by-line PDF measurement of
  // corpora 02/10) shows line heights are NOT rounded to linePitch — 12pt with line
  // spacing 276 (1.15×) lays out at 18pt/line, not the 15.6 or 31.2 grid multiples;
  // space-after is also left unrounded. The earlier round-up was tuned to LibreOffice
  // behavior and halved lines-per-page for docGrid documents (7 pages vs Word's 4).
  // The docGrid param stays: other semantics like lines-per-page rulers may still need it.
  void docGrid

  return lineH
}

/**
 * Font family → canvas CSS font-family fallback chain.
 * When a common Word font is missing locally, fall back to metric-compatible
 * open-source fonts (registered in fonts/fonts.css); identical widths keep line
 * breaks aligned with Word and the offline pagination model. CJK families fall
 * back to macOS equivalents (CJK width is always 1em, so this is mostly glyph appearance).
 */
export function cssFontFamily(font: string): string {
  const f = font.toLowerCase()
  const chain = (...families: string[]) =>
    [...new Set(families)].map((x) => `'${x.replace(/'/g, '')}'`).join(',')
  // CJK fallback at chain end (GB2312 subset bundled in fonts.css): no tofu even without system Chinese fonts
  const CJK_SERIF = 'Noto Serif CJK SC'
  const CJK_SANS = 'Noto Sans CJK SC'
  if (f.includes('calibri')) return `${chain(font, 'Carlito', CJK_SANS)},sans-serif`
  if (f.includes('cambria') && !f.includes('math'))
    return `${chain(font, 'Caladea', CJK_SERIF)},serif`
  if (f.includes('times')) return `${chain(font, 'Liberation Serif', CJK_SERIF)},serif`
  if (f === 'arial' || f.startsWith('arial '))
    return `${chain(font, 'Liberation Sans', CJK_SANS)},sans-serif`
  if (f.includes('courier')) return `${chain(font, 'Liberation Mono', CJK_SANS)},monospace`
  // XiaoBiaoSong/ZhongSong (FZXiaoBiaoSong_GBK etc., gov-document title fonts) before the generic SimSun branch
  if (
    f.includes('小标宋') ||
    f.includes('中宋') ||
    f.includes('xiaobiaosong') ||
    f.includes('zhongsong')
  )
    return `${chain(font, 'STZhongsong', 'Songti SC', 'STSong', 'SimSun', CJK_SERIF)},serif`
  if (f.includes('simsun') || f.includes('宋体') || f.includes('nsimsun')) {
    return `${chain(font, 'Songti SC', 'STSong', 'SimSun', CJK_SERIF)},serif`
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
  // Japanese/Korean/Traditional Chinese: fall back within the same script (win/mac family names as mutual backups) so Han glyphs don't render with Simplified forms
  const JA_SANS = ['Yu Gothic', 'Hiragino Sans', 'Meiryo', 'Noto Sans JP']
  const JA_SERIF = ['Yu Mincho', 'Hiragino Mincho ProN', 'MS Mincho', 'Noto Serif JP']
  const KO_SANS = ['Malgun Gothic', 'Apple SD Gothic Neo', 'Noto Sans KR']
  const KO_SERIF = ['Batang', 'AppleMyungjo', 'Noto Serif KR']
  const TC_SANS = ['Microsoft JhengHei', 'PingFang TC', 'Heiti TC', 'Noto Sans TC']
  const TC_SERIF = ['PMingLiU', 'MingLiU', 'Songti TC', 'Noto Serif TC']
  const nfkc = font.normalize('NFKC')
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
    const serif = /batang|바탕|myeongjo|myungjo|명조|gungsuh|궁서/i.test(nfkc)
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
  const latin = cssFontFamily(ascii)
    .split(',')
    .filter((f) => !/noto (sans|serif) cjk/i.test(f) && !/^(serif|sans-serif|monospace)$/.test(f))
  return `${latin.join(',')},${cssFontFamily(eastAsia)}`
}

/** Text contains CJK characters (decides the line-height factor: CJK lines measure ~1.3em per Chinese font metrics) */
export function textHasCjk(text: string): boolean {
  for (const ch of text) {
    if (isCjk(ch.codePointAt(0) ?? 0)) return true
  }
  return false
}

/**
 * Word line spacing → editing-canvas CSS line-height value (shared by extensions/docStyleCss).
 *
 * Line rules: auto = multiple × font natural line height; atLeast = max(natural, N); exact = N.
 * CSS expression: natural line height ≈ var(--doc-line-factor) (line-height factor
 * of the document's default Chinese font, injected by docStyleCss; default 1.2) × 1em.
 * The canvas previously used a uniform lineSpacing×1.2, which made Chinese documents
 * (factor 1.3) systematically too tight and dropped atLeast/exact entirely.
 */
export function cssLineHeight(
  lineRule: 'auto' | 'atLeast' | 'exact' | undefined,
  lineRawTwips: number | undefined,
  lineSpacing: number | undefined,
): string | null {
  const FACTOR = 'var(--doc-line-factor,1.2)'
  if (lineRule === 'exact' && lineRawTwips) return `${(lineRawTwips / 20).toFixed(1)}pt`
  if (lineRule === 'atLeast' && lineRawTwips) {
    return `max(${(lineRawTwips / 20).toFixed(1)}pt, calc(${FACTOR} * 1em))`
  }
  const m = lineSpacing ?? (lineRule === 'auto' && lineRawTwips ? lineRawTwips / 240 : undefined)
  if (m) return `calc(${FACTOR} * ${m})`
  return null
}

/**
 * Space-before/space-after, aligned to the grid when docGrid exists.
 */
export function snapSpacingToGrid(spacingTwips: number, docGrid: DocGrid | undefined): number {
  // real Word baseline measurement: space-before/after are not rounded to the grid (corpus 10-cn: 240tw space-after lays out at ~8pt, the original value)
  void docGrid
  return spacingTwips * TWIPS_TO_PX
}

// ─── CJK detection ───────────────────────────────────────────────────────────

function isCjk(cp: number): boolean {
  return (
    (cp >= 0x3000 && cp <= 0x30ff) ||
    (cp >= 0x3400 && cp <= 0x9fff) ||
    (cp >= 0xf900 && cp <= 0xfaff) ||
    (cp >= 0xff00 && cp <= 0xffef)
  )
}

// ─── Line-break simulation ───────────────────────────────────────────────────

/**
 * CJK character line-height factor (by font family).
 * PMingLiU/MingLiU (compact Traditional) → 1.0: the font's lineHeight is already 1.0×, no extra boost.
 * Other fonts (SimSun/SimHei/Calibri etc.) → 1.3: CJK chars render via SimSun fallback with ~1.3em line height.
 */
function cjkLineHFactor(fontFamily: string): number {
  const f = fontFamily.toLowerCase()
  if (f.includes('pmingliu') || f.includes('mingliu')) return 1.0
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
    const flushWord = () => {
      if (!buf) return
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
            pushLine(lineH)
            curLineH = lineH
            fragment = ch
          } else {
            fragment += ch
          }
        }
        if (fragment) {
          curLineW += metrics.measure(fragment, style)
          curLineH = Math.max(curLineH, lineH)
          curText = fragment
        }
      } else {
        curLineW += w
        curLineH = Math.max(curLineH, lineH)
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
        flushWord()
        const cw = metrics.measure(ch, style)
        if (curLineW + cw > availWidthPx && curLineW > 0) {
          pushLine(Math.max(curLineH, lineH))
          curLineH = lineH
        }
        curLineW += cw
        curText += ch
        // CJK char line height: NaN = picked dynamically per font family (body path), a number = fixed factor (table cells)
        const runCjkFactor = isNaN(cjkFactor) ? cjkLineHFactor(style.fontFamily) : cjkFactor
        const cjkH = style.fontSizePx * runCjkFactor
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
