/**
 * 2.3 Text layout — replicates PPT text-box layout: wrapping / line spacing /
 * paragraph spacing / vertical alignment / autofit.
 *
 * Input: TextBody (pptx-engine parse result) + text box pixel size + font metrics.
 * Output: RenderTextLayout (glyph runs per line with exact px coordinates).
 *
 * Wrapping rules:
 *   - wrap=true: wrap to the available width; prefer breaking at whitespace, CJK may
 *     break per character, over-long words without whitespace hard-break.
 *   - wrap=false: no wrapping (overflow handled by clipping/autofit).
 *   - autofit=shrink: when content is too tall, shrink font sizes proportionally
 *     (bisection) until it fits or hits the floor.
 *
 * Alignment (baked into glyph coordinates; the renderer computes nothing):
 *   - Paragraph align (l/ctr/r/just) → each line offsets horizontally within the
 *     available width; negative offsets allowed on overflow (wrap=none centered
 *     content overflows equally to both sides).
 *   - bodyPr anchor (t/ctr/b) → the overall vertical offset is added to each line's
 *     top/baselineY.
 */
import type { TextBody, Paragraph, TextRun } from '@genoffice/pptx-engine'
import type { RenderTextLayout, TextLine, GlyphRun } from './render-tree'
import bidiFactory from 'bidi-js'
import { graphemes, isWideChar, type FontMetricsProvider, type RunStyle } from './metrics'
import { emuToPx, ptToPx, type Viewport } from './coords'

const DEFAULT_FONT = 'Arial'
const DEFAULT_SIZE_PT = 18

/** <a:bodyPr> inset defaults (EMU): 0.1" left/right, 0.05" top/bottom. Mirrors the
 *  engine's DEFAULT_BODY_INSETS — importing it would drag the engine runtime into the
 *  renderer bundle, so a test asserts the two stay equal instead. */
export const DEFAULT_INSETS_EMU = { l: 91440, t: 45720, r: 91440, b: 45720 }

interface Token {
  text: string
  style: RunStyle
  color: string
  underline: boolean
  strike?: boolean
  /** Text highlight color (<a:rPr><a:highlight>, drawn as a background behind the run) */
  highlight?: string
  /** Raw super/subscript baseline % (for editor round-trips; blShift is the baked px) */
  blPct?: number
  /** Letter spacing (px/char, may be negative; <a:rPr spc>, counted in both width and rendering) */
  ls: number
  /** Whether a line break is allowed after this token */
  breakable: boolean
  /** Whether it is whitespace (trailing whitespace may be swallowed when wrapping) */
  isSpace: boolean
  /** Tab character: width comes from the paragraph's tab stops, not the font */
  isTab?: boolean
  /** Layout-computed width (px) overriding the font measurement (tab advance) */
  wOverride?: number
  /** Forced line break (<a:br/> soft-break sentinel "\n") */
  isBreak?: boolean
  /** Source model run index (index into Paragraph.runs) */
  srcRun: number
  /** The model run's raw font name (absent = no explicit font; style.fontFamily is a fallback) */
  srcFont?: string
  /** Logical token order before bidi visual reordering (used by editable DOM to restore source order). */
  logicalOrder?: number
  /** UAX#9 direction level (odd = RTL); unset for paragraphs without RTL text */
  level?: number
  /** Super/subscript baseline shift (px, positive = up; <a:rPr baseline> percentage × font size) */
  blShift?: number
  /** Text outline (<a:rPr><a:ln>, commonly used by WordArt) */
  outline?: { color: string; widthPx: number }
  /** Run outer shadow (px, offset resolved from dist/dir) */
  shadow?: { color: string; blurPx: number; offsetX: number; offsetY: number }
  /** WordArt gradient text fill (resolved stops + screen angle) */
  gradient?: { stops: Array<{ pos: number; color: string }>; angleDeg: number; scaled?: boolean }
  /** Run glow (px radius) */
  glow?: { color: string; blurPx: number }
  /** Run reflection (faded mirror below the baseline) */
  reflection?: boolean
  /** Run hyperlink (TextRun.hyperlink encoding) */
  link?: string
}

/** Alpha of a #RRGGBB(AA) color as 0..1 (1 when opaque or unset). */
function hexAlpha(hex: string | undefined): number {
  return hex && hex.length === 9 ? parseInt(hex.slice(7), 16) / 255 : 1
}

/** Multiply a #RRGGBB(AA) color's alpha by factor → #RRGGBBAA. */
function scaleHexAlpha(hex: string, factor: number): string {
  if (factor >= 1) return hex
  const a = Math.round(hexAlpha(hex) * factor * 255)
  return hex.slice(0, 7) + Math.max(0, Math.min(255, a)).toString(16).padStart(2, '0').toUpperCase()
}

function runStyle(run: TextRun, scale: number, fontScale: number): RunStyle {
  const sizePt = run.fontSize ?? DEFAULT_SIZE_PT
  // PowerPoint renders autofit text at round(size × fontScale) whole points — glyphs
  // and the 1.2em line pitch both quantize (probe-measured: 20pt at 46/44/42.5% all
  // draw 9pt, 47.5/48% draw 10pt, 28pt×46%=12.88 draws 13pt, 20pt×52%=10.4 draws 10pt).
  // Explicit fractional sizes without autofit (10.5pt) are untouched.
  const effPt = fontScale !== 1 ? Math.max(1, Math.round(sizePt * fontScale)) : sizePt
  // PowerPoint kerns only at effective size ≥ the rPr kern threshold (absent = 12 pt
  // default, 0 = never; probe-measured). The autofit fontScale counts toward the size.
  const kernMinPt = run.kern ?? 12
  return {
    fontFamily: run.fontFamily || DEFAULT_FONT,
    ...(run.fontScriptHint != null ? { substScript: run.fontScriptHint } : {}),
    fontSizePx: ptToPx(effPt, scale),
    bold: !!run.bold,
    italic: !!run.italic,
    kerning: kernMinPt > 0 && effPt >= kernMinPt,
  }
}

/**
 * PowerPoint never kerns text drawn with a substituted font: an overlapped pair of
 * identical runs (kern default vs kern=0) diverges when the font is installed but
 * coincides pixel-exactly when it's missing (probe-measured). A substituted token
 * measures and draws unkerned, keeping the two sides consistent either way.
 */
function substituteKerning(tok: Token, metrics: FontMetricsProvider): Token {
  if (tok.style.kerning === false || !metrics.substituted?.(tok.style)) return tok
  return { ...tok, style: { ...tok.style, kerning: false } }
}

/** Token width = font advance width + letter spacing × char count (matches canvas letterSpacing: appended after each char) */
function tokenWidth(tok: Token, metrics: FontMetricsProvider): number {
  if (tok.wOverride != null) return tok.wOverride
  const w = metrics.measure(tok.text, tok.style)
  return tok.ls ? w + tok.ls * [...tok.text].length : w
}

/**
 * rPr cap="all"/"small" display transform. Length-preserving: code points whose uppercase
 * expands (ß→SS) stay as-is so click-to-caret offsets keep mapping 1:1 to source text.
 */
function applyCap(text: string, cap: string | undefined): string {
  if (cap !== 'all' && cap !== 'small') return text
  let out = ''
  for (const ch of text) {
    const u = ch.toUpperCase()
    out += [...u].length === 1 ? u : ch
  }
  return out
}

/**
 * Adobe Symbol encoding → Unicode (the differing positions; digits/punctuation are identity).
 * Symbol-font runs store legacy byte codes ('m' = μ): PowerPoint draws them through the
 * font's symbol cmap, but macOS's Symbol.ttf (and every fallback) is Unicode-encoded, so both
 * measuring and drawing need the real code points. Exotic bracket-piece glyphs stay unmapped.
 */
const SYMBOL_TO_UNICODE: Record<number, number> = {
  0x22: 0x2200,
  0x24: 0x2203,
  0x27: 0x220b,
  0x2a: 0x2217,
  0x2d: 0x2212,
  0x40: 0x2245,
  0x41: 0x0391,
  0x42: 0x0392,
  0x43: 0x03a7,
  0x44: 0x0394,
  0x45: 0x0395,
  0x46: 0x03a6,
  0x47: 0x0393,
  0x48: 0x0397,
  0x49: 0x0399,
  0x4a: 0x03d1,
  0x4b: 0x039a,
  0x4c: 0x039b,
  0x4d: 0x039c,
  0x4e: 0x039d,
  0x4f: 0x039f,
  0x50: 0x03a0,
  0x51: 0x0398,
  0x52: 0x03a1,
  0x53: 0x03a3,
  0x54: 0x03a4,
  0x55: 0x03a5,
  0x56: 0x03c2,
  0x57: 0x03a9,
  0x58: 0x039e,
  0x59: 0x03a8,
  0x5a: 0x0396,
  0x5c: 0x2234,
  0x5e: 0x22a5,
  0x60: 0x203e,
  0x61: 0x03b1,
  0x62: 0x03b2,
  0x63: 0x03c7,
  0x64: 0x03b4,
  0x65: 0x03b5,
  0x66: 0x03c6,
  0x67: 0x03b3,
  0x68: 0x03b7,
  0x69: 0x03b9,
  0x6a: 0x03d5,
  0x6b: 0x03ba,
  0x6c: 0x03bb,
  0x6d: 0x03bc,
  0x6e: 0x03bd,
  0x6f: 0x03bf,
  0x70: 0x03c0,
  0x71: 0x03b8,
  0x72: 0x03c1,
  0x73: 0x03c3,
  0x74: 0x03c4,
  0x75: 0x03c5,
  0x76: 0x03d6,
  0x77: 0x03c9,
  0x78: 0x03be,
  0x79: 0x03c8,
  0x7a: 0x03b6,
  0x7e: 0x223c,
  0xa1: 0x03d2,
  0xa2: 0x2032,
  0xa3: 0x2264,
  0xa4: 0x2044,
  0xa5: 0x221e,
  0xa6: 0x0192,
  0xa7: 0x2663,
  0xa8: 0x2666,
  0xa9: 0x2665,
  0xaa: 0x2660,
  0xab: 0x2194,
  0xac: 0x2190,
  0xad: 0x2191,
  0xae: 0x2192,
  0xaf: 0x2193,
  0xb0: 0x00b0,
  0xb1: 0x00b1,
  0xb2: 0x2033,
  0xb3: 0x2265,
  0xb4: 0x00d7,
  0xb5: 0x221d,
  0xb6: 0x2202,
  0xb7: 0x2022,
  0xb8: 0x00f7,
  0xb9: 0x2260,
  0xba: 0x2261,
  0xbb: 0x2248,
  0xbc: 0x2026,
  0xbf: 0x21b5,
  0xc0: 0x2135,
  0xc1: 0x2111,
  0xc2: 0x211c,
  0xc3: 0x2118,
  0xc4: 0x2297,
  0xc5: 0x2295,
  0xc6: 0x2205,
  0xc7: 0x2229,
  0xc8: 0x222a,
  0xc9: 0x2283,
  0xca: 0x2287,
  0xcb: 0x2284,
  0xcc: 0x2282,
  0xcd: 0x2286,
  0xce: 0x2208,
  0xcf: 0x2209,
  0xd0: 0x2220,
  0xd1: 0x2207,
  0xd2: 0x00ae,
  0xd3: 0x00a9,
  0xd4: 0x2122,
  0xd5: 0x220f,
  0xd6: 0x221a,
  0xd7: 0x22c5,
  0xd8: 0x00ac,
  0xd9: 0x2227,
  0xda: 0x2228,
  0xdb: 0x21d4,
  0xdc: 0x21d0,
  0xdd: 0x21d1,
  0xde: 0x21d2,
  0xdf: 0x21d3,
  0xe0: 0x25ca,
  0xe1: 0x2329,
  0xe5: 0x2211,
  0xf1: 0x232a,
  0xf2: 0x222b,
}
const SYMBOL_FONT_RE = /^symbol$/i
/** Translate a Symbol-encoded run to Unicode (accepts both raw bytes and F0xx PUA storage). */
function symbolRunText(text: string): string {
  let out = ''
  for (const ch of text) {
    let cp = ch.codePointAt(0) ?? 0
    if (cp >= 0xf000 && cp <= 0xf0ff) cp -= 0xf000
    // Unmapped positions: ASCII identity slots (digits/punctuation) use the stripped code
    // even when stored as PUA; high unmapped slots (bracket pieces) keep the original char
    // (stripping would fabricate Latin-1 letters like U+00E6).
    const fallback = cp >= 0x20 && cp <= 0x7f ? cp : (ch.codePointAt(0) ?? 0)
    out += String.fromCodePoint(SYMBOL_TO_UNICODE[cp] ?? fallback)
  }
  return out
}

/** Splits a paragraph's runs into a breakable token stream (whitespace / single CJK chars / Latin words). */
function tokenizeParagraph(p: Paragraph, scale: number, fontScale: number): Token[] {
  const tokens: Token[] = []
  p.runs.forEach((run, srcRun) => {
    const style = runStyle(run, scale, fontScale)
    const color = run.color ?? '#000000'
    const underline = !!run.underline
    const ls = run.letterSpacing ? ptToPx(run.letterSpacing, scale) * fontScale : 0
    // Super/subscript: baseline% (30 = superscript raised 30% of font size, negative = subscript lowered)
    const blShift = run.baseline ? style.fontSizePx * (run.baseline / 100) : 0
    const base = {
      style,
      color,
      underline,
      ls,
      srcRun,
      ...(run.fontFamily ? { srcFont: run.fontFamily } : {}),
      ...(run.hyperlink ? { link: run.hyperlink } : {}),
      ...(run.strike ? { strike: true } : {}),
      ...(run.highlight ? { highlight: run.highlight } : {}),
      ...(run.baseline ? { blPct: run.baseline } : {}),
      ...(blShift ? { blShift } : {}),
      ...(run.outline
        ? {
            outline: {
              color: run.outline.color,
              widthPx: emuToPx(run.outline.widthEmu, scale) * fontScale,
            },
          }
        : {}),
      ...(run.gradient
        ? {
            gradient: {
              stops: run.gradient.stops.map((s) => ({ pos: s.pos, color: s.color })),
              // OOXML gradient angle is 1/60000° clockwise from the +x axis
              angleDeg: (run.gradient.angle ?? 5400000) / 60000,
              ...(run.gradient.scaled ? { scaled: true } : {}),
            },
          }
        : {}),
      ...(run.glow
        ? {
            glow: {
              color: run.glow.color,
              blurPx: emuToPx(run.glow.radius, scale) * fontScale,
            },
          }
        : {}),
      ...(run.reflection ? { reflection: true } : {}),
      ...(run.shadow
        ? {
            shadow: {
              // A translucent fill dims its own shadow (PowerPoint-observed): scale the
              // shadow alpha by the text fill alpha
              color: scaleHexAlpha(run.shadow.color, hexAlpha(run.color)),
              blurPx: emuToPx(run.shadow.blurRad, scale) * fontScale,
              offsetX:
                emuToPx(run.shadow.dist, scale) *
                fontScale *
                Math.cos((run.shadow.dirDeg * Math.PI) / 180),
              offsetY:
                emuToPx(run.shadow.dist, scale) *
                fontScale *
                Math.sin((run.shadow.dirDeg * Math.PI) / 180),
            },
          }
        : {}),
    }
    let buf = ''
    const flushWord = () => {
      if (!buf) return
      if (WORD_SEG && SEA_RE.test(buf)) {
        // Southeast Asian scripts without spaces: ICU dictionary segmentation; word gaps are break opportunities
        for (const s of WORD_SEG.segment(buf)) {
          tokens.push({ ...base, text: s.segment, breakable: true, isSpace: false })
        }
      } else {
        tokens.push({ ...base, text: buf, breakable: false, isSpace: false })
      }
      buf = ''
    }
    // Symbol-encoded runs become Unicode before cap/segmentation (cap would uppercase μ→Μ)
    const runText = SYMBOL_FONT_RE.test(style.fontFamily)
      ? symbolRunText(run.text)
      : applyCap(run.text, run.cap)
    // Per grapheme cluster: combining marks / ZWJ emoji always share a token with their base char, never split by wrapping
    for (const ch of graphemes(runText)) {
      const cp = ch.codePointAt(0) ?? 0
      if (ch === '\n' || ch === '\v') {
        flushWord()
        tokens.push({ ...base, text: '\n', breakable: true, isSpace: false, isBreak: true })
      } else if (ch === ' ' || ch === '　') {
        // U+3000 ideographic space wraps like whitespace (PowerPoint swallows it at line ends)
        flushWord()
        tokens.push({ ...base, text: ch, breakable: true, isSpace: true })
      } else if (ch === '\t') {
        flushWord()
        tokens.push({ ...base, text: ch, breakable: true, isSpace: true, isTab: true })
      } else if (ch === ' ') {
        // NBSP: glues neighbors into one token but draws/measures at plain-space
        // width (fonts like Carlito have no U+00A0 glyph → the missing-glyph
        // fallback would badly over-measure it)
        buf += ' '
      } else if (isWideChar(cp)) {
        flushWord()
        tokens.push({ ...base, text: ch, breakable: true, isSpace: false })
      } else if (BREAK_AFTER_DASH.has(cp) && buf) {
        // Hyphen/dash after word characters is a break-after opportunity (PowerPoint
        // wraps "Cloud–Edge" as "Cloud–" + "Edge", never mid-word)
        buf += ch
        flushWord()
      } else {
        buf += ch
      }
    }
    flushWord()
  })
  return tokens
}

function toRoman(n: number): string {
  const table: Array<[number, string]> = [
    [1000, 'm'],
    [900, 'cm'],
    [500, 'd'],
    [400, 'cd'],
    [100, 'c'],
    [90, 'xc'],
    [50, 'l'],
    [40, 'xl'],
    [10, 'x'],
    [9, 'ix'],
    [5, 'v'],
    [4, 'iv'],
    [1, 'i'],
  ]
  let out = ''
  for (const [v, s] of table)
    while (n >= v) {
      out += s
      n -= v
    }
  return out
}

function toAlpha(n: number): string {
  let out = ''
  while (n > 0) {
    n--
    out = String.fromCharCode(97 + (n % 26)) + out
    n = Math.floor(n / 26)
  }
  return out
}

const CJK_DIGITS = '〇一二三四五六七八九'
function toCjkNum(n: number): string {
  if (n <= 10) return n === 10 ? '十' : CJK_DIGITS[n]!
  if (n < 20) return '十' + CJK_DIGITS[n % 10]!
  if (n < 100) return CJK_DIGITS[Math.floor(n / 10)]! + '十' + (n % 10 ? CJK_DIGITS[n % 10]! : '')
  return String(n)
}

/**
 * <a:buAutoNum type> → numbered-bullet glyph (ST_TextAutonumberScheme). Circled numbers
 * only exist up to ⑳/⓴; PowerPoint falls back to plain arabic beyond that.
 */
function formatAutoNum(n: number, numType: string | undefined): string {
  const t = numType ?? 'arabicPeriod'
  if (t.startsWith('circleNum')) {
    if (t === 'circleNumWdBlackPlain')
      return n <= 10
        ? String.fromCodePoint(0x2775 + n) // ❶–❿
        : n <= 20
          ? String.fromCodePoint(0x24eb + (n - 11)) // ⓫–⓴
          : String(n)
    // circleNumWdWhitePlain included: Wingdings white circled digits are single-ring, i.e. ①–⑳
    return n <= 20 ? String.fromCodePoint(0x245f + n) : String(n) // ①–⑳
  }
  let body: string
  if (t.startsWith('alphaLc')) body = toAlpha(n)
  else if (t.startsWith('alphaUc')) body = toAlpha(n).toUpperCase()
  else if (t.startsWith('romanLc')) body = toRoman(n)
  else if (t.startsWith('romanUc')) body = toRoman(n).toUpperCase()
  else if (t.startsWith('arabicDb'))
    body = [...String(n)].map((d) => String.fromCodePoint(0xff10 + Number(d))).join('') // fullwidth １２３
  else if (t.startsWith('ea1Chs') || t.startsWith('ea1Cht')) body = toCjkNum(n)
  else body = String(n)
  if (t.endsWith('ParenBoth')) return `(${body})`
  if (t.endsWith('ParenR')) return `${body})`
  if (t.endsWith('Period')) return `${body}.`
  if (t.endsWith('Plain')) return body
  return `${body}.`
}

/**
 * Legacy symbol-font bullets (<a:buFont> Wingdings/Webdings): the glyph must be drawn
 * with that font, and ASCII bullet codes normalized into the font's F0xx PUA range
 * (files carry either encoding; the shipped fonts map the PUA).
 */
const SYMBOL_BULLET_RE = /^(wingdings|webdings)/i
function symbolBulletText(font: string | undefined, char: string): string | undefined {
  if (!font || !SYMBOL_BULLET_RE.test(font)) return undefined
  const cp = char.codePointAt(0) ?? 0
  if (cp >= 0xf000 && cp <= 0xf0ff) return char
  return cp >= 0x20 && cp <= 0xff ? String.fromCodePoint(0xf000 + cp) : char
}

// ── East-Asian kinsoku (PowerPoint default rules) ───────────────────
// Closing marks / small kana must not begin a line; opening brackets must not end one.
// Only single-grapheme tokens participate (CJK tokenization emits one token per wide char).
const KINSOKU_NO_START = new Set(
  '、。々,.!?:;)]}%，．！？：；）］｝％〉》」』】〕〟｡｣､･ｰﾞﾟ・ーぁぃぅぇぉっゃゅょゎゕゖァィゥェォッャュョヮヵヶ゛゜ゝゞヽヾ',
)
const KINSOKU_NO_END = new Set('([{$（［｛＄〈《「『【〔〝｢£¥￡￥')
const kinsokuNoStart = (t: Token) => KINSOKU_NO_START.has(t.text)
const kinsokuNoEnd = (t: Token) => KINSOKU_NO_END.has(t.text)

/** Dashes that allow a break after them (U+2011 non-breaking hyphen intentionally absent). */
const BREAK_AFTER_DASH = new Set([0x2d, 0x2010, 0x2012, 0x2013, 0x2014])

/** Southeast Asian scripts without spaces (Thai/Burmese/Khmer/Lao); word boundaries need ICU dictionary segmentation. */
const SEA_RE = /[฀-໿က-႟ក-៿]/
const WORD_SEG: Intl.Segmenter | null =
  typeof Intl !== 'undefined' && 'Segmenter' in Intl
    ? new Intl.Segmenter(undefined, { granularity: 'word' })
    : null

// ── Bidi (UAX#9) ────────────────────────────────────────────────────

/** Strong RTL characters (Hebrew/Arabic + presentation forms) — bidi analysis runs only on a match. */
const RTL_RE = /[\u0590-\u08ff\ufb1d-\ufdff\ufe70-\ufeff]/
let bidiApi: ReturnType<typeof bidiFactory> | null = null

/**
 * Runs UAX#9 over the paragraph token stream: tokens crossing direction levels are
 * split at level boundaries and tagged with their level. Line breaking stays in
 * logical order (UAX#9 requires wrap first, reorder after); visualOrder reorders
 * each line once formed.
 */
function applyBidi(tokens: Token[], baseRtl?: boolean): Token[] {
  const text = tokens.map((t) => t.text).join('')
  // An explicit RTL base still needs analysis for pure-LTR text (neutrals move to the left edge)
  if (!RTL_RE.test(text) && !baseRtl) return tokens
  bidiApi ??= bidiFactory()
  const { levels } = bidiApi.getEmbeddingLevels(
    text,
    baseRtl == null ? undefined : baseRtl ? 'rtl' : 'ltr',
  )
  const out: Token[] = []
  let off = 0
  for (const tok of tokens) {
    const end = off + tok.text.length
    let segStart = off
    for (let i = off + 1; i <= end; i++) {
      if (i === end || levels[i] !== levels[segStart]) {
        out.push({ ...tok, text: text.slice(segStart, i), level: levels[segStart] ?? 0 })
        segStart = i
      }
    }
    off = end
  }
  return out
}

/** UAX#9 L2: reorder line tokens into visual order (from the highest level down to the lowest odd level, reversing contiguous runs per level). */
function visualOrder(toks: Token[]): Token[] {
  let max = 0
  let minOdd = Infinity
  for (const t of toks) {
    const lv = t.level ?? 0
    if (lv > max) max = lv
    if (lv % 2 === 1 && lv < minOdd) minOdd = lv
  }
  if (minOdd === Infinity) return toks
  const arr = [...toks]
  for (let l = max; l >= minOdd; l--) {
    for (let i = 0; i < arr.length;) {
      if ((arr[i]!.level ?? 0) >= l) {
        let j = i
        while (j < arr.length && (arr[j]!.level ?? 0) >= l) j++
        for (let a = i, b = j - 1; a < b; a++, b--) {
          const t = arr[a]!
          arr[a] = arr[b]!
          arr[b] = t
        }
        i = j
      } else {
        i++
      }
    }
  }
  return arr
}

/** Paragraph base direction: explicit a:pPr rtl wins, else the first strong-direction character (used for the default alignment; RTL paragraphs without explicit alignment align right). */
function paraBaseRtl(p: Paragraph): boolean {
  if (p.rtl != null) return p.rtl
  for (const r of p.runs) {
    for (const ch of r.text) {
      if (RTL_RE.test(ch)) return true
      // eslint-disable-next-line no-misleading-character-class -- broad strong-LTR ranges; combining marks inside are irrelevant for a per-char strong-direction probe
      if (/[A-Za-z\u00c0-\u058f\u0900-\ud7ff\uf900-\ufdcf]/.test(ch)) return false
    }
  }
  return false
}

interface LaidLine {
  runs: GlyphRun[]
  height: number
  ascent: number
  descent: number
  /** Single line height (without lnSpc multiplier/reduction; the basis for spcPct paragraph spacing) */
  singleH: number
  /** Baseline offset minus the line's max ascent (signed): baseline = top + leadAbove + ascent.
   *  Positive when the box is taller than the glyphs (lnSpc>100%), negative when shorter (lnSpc<100%
   *  or fonts whose ascent+descent exceed the 1.2em box). */
  leadAbove?: number
  /** Trailing whitespace was swallowed when wrapping (tells the editor to re-add a space when joining lines) */
  trailingSpace?: boolean
  /** The exact swallowed whitespace (may be U+3000 / tabs; the editor re-adds it verbatim) */
  trailingText?: string
  /** The line ends at a soft-break sentinel; value = the sentinel run's model index */
  softBreakAfter?: number
}

/** Wraps a single paragraph → array of lines (coordinates relative to the x origin; y added uniformly later). */
function layoutParagraph(
  p: Paragraph,
  availWidth: number,
  wrap: boolean,
  metrics: FontMetricsProvider,
  scale: number,
  fontScale: number,
  lnSpcRed = 0,
  /** First-line x shift (px): the first-line indent, or the push of a bullet glyph
   *  overflowing the hanging indent. Shrinks (negative: widens) the first line's wrap
   *  budget and offsets its tab cursor by the same amount. */
  firstLineShrinkPx = 0,
  /** Tab geometry: stops + default grid in px measured from the text-frame left inset,
   *  originPx = the line's x offset (marL) in that same space. */
  tabs?: { stopsPx: number[]; defaultPx: number; originPx: number },
): LaidLine[] {
  const tokens = applyBidi(tokenizeParagraph(p, scale, fontScale), p.rtl).map(
    (tok, logicalOrder) => ({
      ...substituteKerning(tok, metrics),
      logicalOrder,
    }),
  )
  const lines: LaidLine[] = []
  let cur: Token[] = []
  let curW = 0

  const pushLine = (toks: Token[]) => {
    // Strip trailing whitespace (recorded verbatim: the editor re-adds it when joining wrapped lines back into a paragraph)
    let trailingSpace = false
    let trailingText = ''
    while (toks.length && toks[toks.length - 1]!.isSpace) {
      trailingText = toks.pop()!.text + trailingText
      trailingSpace = true
    }
    if (!toks.length) {
      // An empty line still occupies one line height. A textless run (the marker
      // element-level font ops leave on blank cells/paragraphs) styles the line and is
      // kept as a zero-width glyph run so the editor and ribbon readback see the format.
      const src = p.runs[0]
      const st: RunStyle = src
        ? runStyle(src, scale, fontScale)
        : {
            fontFamily: DEFAULT_FONT,
            fontSizePx: ptToPx(
              fontScale !== 1
                ? Math.max(1, Math.round(DEFAULT_SIZE_PT * fontScale))
                : DEFAULT_SIZE_PT,
              scale,
            ),
            bold: false,
            italic: false,
          }
      const m = metrics.metrics(st)
      const box = lineH(p, st.fontSizePx, scale, lnSpcRed)
      const leadAbove = baselineOff(p, box, m.descent) - m.ascent
      lines.push({
        runs: src
          ? [
              {
                text: '',
                x: 0,
                baselineY: 0,
                fontFamily: st.fontFamily,
                ...(src.fontFamily ? { srcFontFamily: src.fontFamily } : {}),
                fontSizePx: st.fontSizePx,
                color: src.color ?? '#000000',
                bold: st.bold,
                italic: st.italic,
                underline: !!src.underline,
                widthPx: 0,
                srcRunIdx: 0,
                ascentPx: m.ascent,
              },
            ]
          : [],
        height: box,
        ...(leadAbove ? { leadAbove } : {}),
        ascent: m.ascent,
        descent: m.descent,
        singleH: PPT_SINGLE * st.fontSizePx,
        ...(trailingSpace ? { trailingSpace, trailingText } : {}),
      })
      return
    }
    const line = buildLine(toks, metrics, p, scale, lnSpcRed)
    if (trailingSpace) {
      line.trailingSpace = true
      line.trailingText = trailingText
    }
    lines.push(line)
  }

  let endedWithBreak = false
  for (const tok of tokens) {
    // <a:br/> forced break: breaks regardless of wrap; record the sentinel run index for editor round-trips
    if (tok.isBreak) {
      pushLine(cur)
      lines[lines.length - 1]!.softBreakAfter = tok.srcRun
      cur = []
      curW = 0
      endedWithBreak = true
      continue
    }
    endedWithBreak = false
    if (tok.isTab && tabs) {
      // Tab advances to the next stop past the cursor; past the last stop, to the default grid
      const cursor = tabs.originPx + (lines.length === 0 ? firstLineShrinkPx : 0) + curW
      const stop = tabs.stopsPx.find((s) => s > cursor + 0.5)
      tok.wOverride = Math.max(
        (stop ?? (Math.floor(cursor / tabs.defaultPx) + 1) * tabs.defaultPx) - cursor,
        0,
      )
    }
    const w = tokenWidth(tok, metrics)
    // The first line loses firstLineShrinkPx to the overflowing bullet glyph; evaluated
    // lazily because the soft wrap right below can end line 0 for this same token
    const lineAvail = () => (lines.length === 0 ? availWidth - firstLineShrinkPx : availWidth)
    if (wrap && cur.length && curW + w > lineAvail() && !tok.isSpace) {
      // Kinsoku: pull the predecessor down when the new line would start with a closing
      // mark, push an opening bracket down when it would end the old line.
      const carry: Token[] = []
      while (cur.length > 1) {
        const head = carry[0] ?? tok
        const last = cur[cur.length - 1]!
        if (last.isSpace || (!kinsokuNoStart(head) && !kinsokuNoEnd(last))) break
        carry.unshift(cur.pop()!)
      }
      pushLine(cur)
      cur = carry
      curW = carry.reduce((s, t) => s + tokenWidth(t, metrics), 0)
    }
    // Hard-break over-long words (a single token wider than the line)
    if (wrap && !cur.length && w > lineAvail() && tok.text.length > 1 && !tok.isSpace) {
      for (const seg of hardBreak(tok, lineAvail(), metrics)) {
        pushLine([seg])
      }
      continue
    }
    cur.push(tok)
    curW += w
  }
  if (cur.length) pushLine(cur)
  // A trailing <a:br/> shows an empty last line (PowerPoint counts it when anchoring)
  else if (endedWithBreak) pushLine([])
  if (!lines.length) pushLine([])
  return lines
}

function hardBreak(tok: Token, availWidth: number, metrics: FontMetricsProvider): Token[] {
  const out: Token[] = []
  let buf = ''
  // Hard-cut per grapheme cluster: cut points never land inside combining-mark/ZWJ sequences
  for (const ch of graphemes(tok.text)) {
    const test = buf + ch
    if (buf && tokenWidth({ ...tok, text: test }, metrics) > availWidth) {
      out.push({ ...tok, text: buf })
      buf = ch
    } else {
      buf = test
    }
  }
  if (buf) out.push({ ...tok, text: buf })
  return out
}

function buildLine(
  toks: Token[],
  metrics: FontMetricsProvider,
  p: Paragraph,
  scale: number,
  lnSpcRed = 0,
): LaidLine {
  let x = 0
  let ascent = 0
  let descent = 0
  let sizeM = 0
  const runs: GlyphRun[] = []
  // Lines with RTL levels are reordered into visual order before laying out x (break in logical order → draw in visual order)
  for (const tok of visualOrder(toks)) {
    const m = metrics.metrics(tok.style)
    ascent = Math.max(ascent, m.ascent)
    descent = Math.max(descent, m.descent)
    sizeM = Math.max(sizeM, tok.style.fontSizePx)
    const w = tokenWidth(tok, metrics)
    runs.push({
      text: tok.text,
      x,
      baselineY: 0, // filled in later from the line's ascent
      // When a missing font is substituted, draw with the substitute name so drawing and measuring use the same font file
      fontFamily: metrics.displayFamily?.(tok.style, tok.text) ?? tok.style.fontFamily,
      ...(tok.srcFont ? { srcFontFamily: tok.srcFont } : {}),
      fontSizePx: tok.style.fontSizePx,
      color: tok.color,
      bold: tok.style.bold,
      italic: tok.style.italic,
      underline: tok.underline,
      ...(tok.strike ? { strike: true } : {}),
      ...(tok.highlight ? { highlight: tok.highlight } : {}),
      widthPx: w,
      ...(tok.ls ? { letterSpacingPx: tok.ls } : {}),
      ...(tok.style.kerning === false ? { kerningOff: true } : {}),
      ...(tok.outline ? { outline: tok.outline } : {}),
      ...(tok.gradient ? { gradient: tok.gradient } : {}),
      ...(tok.glow ? { glow: tok.glow } : {}),
      ...(tok.reflection ? { reflection: true } : {}),
      ...(tok.shadow ? { shadow: tok.shadow } : {}),
      ...(tok.blShift ? { baselineShiftPx: tok.blShift } : {}),
      ...(tok.blPct ? { baselinePct: tok.blPct } : {}),
      ...(tok.level != null && tok.level % 2 === 1 ? { rtl: true } : {}),
      srcRunIdx: tok.srcRun,
      ...(tok.link ? { link: tok.link } : {}),
      ...(tok.logicalOrder != null ? { logicalOrder: tok.logicalOrder } : {}),
      ascentPx: m.ascent,
    })
    x += w
  }
  const box = lineH(p, sizeM, scale, lnSpcRed)
  const leadAbove = baselineOff(p, box, descent) - ascent
  return {
    runs,
    height: box,
    ...(leadAbove ? { leadAbove } : {}),
    ascent,
    descent,
    singleH: PPT_SINGLE * sizeM,
  }
}

/**
 * PowerPoint line model (mac PPT, 48pt probe deck, 2026-08-16 — font-independent):
 *   - default / explicit 100%: line box = 1.2 × fontSize (regardless of font metrics);
 *     the baseline is BOTTOM-anchored: box − the font's real descent.
 *   - explicit spcPct ≠ 100 (incl. <100) and spcPts: box = pct × 1.2em (or the absolute
 *     points), and the baseline sits at 0.7333 × box (a generic 0.88/0.32 em split scaled
 *     with the box) — real ascent/descent are NOT consulted, so at 70% the ink pokes
 *     above the box top exactly like PowerPoint.
 * All ten probed fonts (incl. MS PGothic 1.0em and Malgun 1.32em hhea) pace at 1.2em.
 */
const PPT_SINGLE = 1.2
const PPT_BASELINE_FRAC = 0.88 / 1.2

function lineH(p: Paragraph, sizePx: number, scale: number, lnSpcRed = 0): number {
  if (p.lineExact != null) return ptToPx(p.lineExact, scale) // exact line spacing is unaffected by autofit spacing reduction
  const single = PPT_SINGLE * sizePx
  const base = p.lineHeight != null ? single * (p.lineHeight / 100) : single
  return base * (1 - lnSpcRed)
}

/** Baseline offset from the line-box top (see the model above). */
function baselineOff(p: Paragraph, box: number, descent: number): number {
  // Explicit spacing (spcPct != 100 and spcPts at every size — 36..120pt probed): the
  // generic 0.7333 split; only default/100% consults the font's real descent.
  if (p.lineExact != null || (p.lineHeight != null && p.lineHeight !== 100)) {
    return PPT_BASELINE_FRAC * box
  }
  return box - descent
}

export interface TextLayoutInput {
  body: TextBody
  boxWidthPx: number
  boxHeightPx: number
  metrics: FontMetricsProvider
  vp: Viewport
  /** Table cells: PowerPoint drops the first paragraph's space-before and the last
      paragraph's space-after (rows stay at their minimum height regardless of them). */
  trimEdgeSpacing?: boolean
  /** Re-run the autofit ladder below a stored fontScale (edit flows only): plain rendering
      honors the cache as-is like PowerPoint on open. */
  refitAutofit?: boolean
}

/**
 * bodyPr numCol: distribute a single laid-out line stream into columns. PowerPoint fills
 * a column to the box height, then continues at the top of the next (no balancing); a line
 * never splits across the boundary. Lines are rebased to the column top and shifted right
 * by whole column strides; the last column keeps any overflow.
 */
function flowIntoColumns(
  result: { lines: TextLine[]; contentHeight: number },
  availHeight: number,
  numCol: number,
  stride: number,
): { lines: TextLine[]; contentHeight: number } {
  const out: TextLine[] = []
  let col = 0
  let colTop = 0 // original-stream top of the current column's first line
  let maxBottom = 0
  for (const ln of result.lines) {
    let top = ln.top - colTop
    if (col < numCol - 1 && top > 0 && top + ln.height > availHeight) {
      col += 1
      colTop = ln.top
      top = 0
    }
    const dx = col * stride
    const dy = top - ln.top
    out.push({
      ...ln,
      top,
      runs:
        dx || dy
          ? ln.runs.map((r) => ({ ...r, x: r.x + dx, baselineY: r.baselineY + dy }))
          : ln.runs,
    })
    maxBottom = Math.max(maxBottom, top + ln.height)
  }
  return { lines: out, contentHeight: maxBottom }
}

/** Main entry: TextBody → RenderTextLayout (including autofit shrink stepping). */
export function layoutText(input: TextLayoutInput): RenderTextLayout {
  const { body, boxWidthPx, boxHeightPx, metrics, vp } = input
  const insets = {
    l: emuToPx(body.insets?.l ?? DEFAULT_INSETS_EMU.l, vp.scale),
    t: emuToPx(body.insets?.t ?? DEFAULT_INSETS_EMU.t, vp.scale),
    r: emuToPx(body.insets?.r ?? DEFAULT_INSETS_EMU.r, vp.scale),
    b: emuToPx(body.insets?.b ?? DEFAULT_INSETS_EMU.b, vp.scale),
  }
  const availWidth = Math.max(boxWidthPx - insets.l - insets.r, 1)
  const availHeight = Math.max(boxHeightPx - insets.t - insets.b, 1)
  const wrap = body.wrap !== false

  // vert/vert270 rotate the whole block (CJK glyphs included); eaVert/wordArtVert
  // are upright column layouts (wordArtVert additionally keeps Latin upright)
  if (body.vert === 'vert' || body.vert === 'vert270') return layoutTextRotated(input, body.vert)
  if (body.vert)
    return {
      ...layoutTextVertical(
        body,
        body.vert,
        availWidth,
        availHeight,
        insets,
        wrap,
        metrics,
        vp.scale,
      ),
      autofit: body.autofit ?? 'none',
    }

  // bodyPr numCol: paragraphs wrap at the column width and fill column after column
  const numCol = body.numCol && body.numCol > 1 ? Math.floor(body.numCol) : 1
  const colGapPx = numCol > 1 ? emuToPx(body.spcCol ?? 0, vp.scale) : 0
  const colWidth =
    numCol > 1 ? Math.max((availWidth - (numCol - 1) * colGapPx) / numCol, 1) : availWidth

  const build = (fontScale: number, lnSpcRed: number) =>
    layoutAll(body, colWidth, wrap, metrics, vp.scale, fontScale, lnSpcRed, input.trimEdgeSpacing)

  // PowerPoint's stored shrink ratio takes priority (files with shrunk text render
  // as-is; we don't scale back up per our own metrics). Only if content still
  // overflows (we edited text / metric differences) do we step further down the
  // discrete autofit ladder (see SHRINK_STEPS), capped at the stored value
  const storedScale = body.autofit === 'shrink' ? (body.fontScale ?? 1) : 1
  const storedRed = body.autofit === 'shrink' ? (body.lnSpcReduction ?? 0) : 0
  let fontScale = storedScale
  let lnSpcReduction = storedRed
  let result = build(fontScale, lnSpcReduction)
  // PowerPoint's autofit ignores blank lines on both ends (files where the stored
  // ratio only fits when they are excluded), so overflow is judged by the span from
  // the first to the last non-blank line — plus 3% slack for residual metric drift,
  // since PowerPoint renders the authored scale as-is.
  const fitSpan = (r: { lines: TextLine[]; contentHeight: number }): number => {
    let top: number | undefined
    let bottom: number | undefined
    for (const ln of r.lines) {
      if (!ln.runs.some((run) => run.text.trim())) continue
      top ??= ln.top
      bottom = ln.top + ln.height
    }
    return bottom === undefined ? r.contentHeight : bottom - (top ?? 0)
  }
  // With columns, the single-stream span may spread over numCol columns of availHeight
  const fitTarget = availHeight * numCol
  // A stored fontScale is authoritative: PowerPoint renders the cached ratio as-is on
  // open/export and never re-fits (probe-measured — a bare <a:normAutofit/> even renders
  // at 100% overflowing the box). Stepping below the cache when our metrics run a hair
  // long shrank whole pages a visible notch (autofit ladder only serves cache-less boxes,
  // i.e. live edits).
  if (
    body.autofit === 'shrink' &&
    (input.refitAutofit || body.fontScale == null) &&
    fitSpan(result) > fitTarget * 1.03
  ) {
    for (const [fs, red] of SHRINK_STEPS) {
      if (fs >= storedScale - 1e-6) continue
      const effRed = Math.max(red, storedRed)
      const r = build(fs, effRed)
      fontScale = fs
      lnSpcReduction = effRed
      result = r
      if (fitSpan(r) <= fitTarget) break
      // Still doesn't fit after all steps → stop at 25% (overflow allowed)
    }
  }
  if (numCol > 1) result = flowIntoColumns(result, availHeight, numCol, colWidth + colGapPx)

  // Vertical anchor: middle/bottom shift everything down (negative offsets allowed
  // when content is too tall → overflow upward/on both sides),
  // baked straight into line coordinates so the renderer draws as-is.
  const anchor = body.anchor ?? 'top'
  // Anchor against the glyph extent: with explicit lnSpc the line box extends below the
  // ink, and rows/boxes sized to the ink (auto table rows) would otherwise get a negative
  // offset that clips the top. Equal to contentHeight under single spacing.
  const extraH = availHeight - (result.inkBottom ?? result.contentHeight)
  const dy = anchor === 'middle' ? extraH / 2 : anchor === 'bottom' ? extraH : 0
  // bodyPr anchorCtr="1": the whole text block (widest-line bounding box) centers
  // horizontally in the text area; paragraph alignment stays relative within the block.
  let dxCtr = 0
  if (body.anchorCtr && numCol === 1) {
    let minX = Infinity
    let maxX = -Infinity
    for (const ln of result.lines) {
      for (const r of ln.runs) {
        if (!r.text) continue
        if (r.x < minX) minX = r.x
        if (r.x + r.widthPx > maxX) maxX = r.x + r.widthPx
      }
    }
    if (maxX > minX) dxCtr = (availWidth - (maxX - minX)) / 2 - minX
  }
  const lines =
    dy || dxCtr
      ? result.lines.map((ln) => ({
          ...ln,
          top: ln.top + dy,
          runs: ln.runs.map((r) => ({ ...r, x: r.x + dxCtr, baselineY: r.baselineY + dy })),
        }))
      : result.lines

  // WordArt text extrusion: depth projected by the body camera tilt (same screen basis
  // as the scene3d shape pipeline: depth leans (sin lon, -sin lat·cos lon) per unit)
  let extrusion: RenderTextLayout['extrusion']
  if (body.extrusion3d) {
    const e = body.extrusion3d
    const d = emuToPx(e.depthEmu, vp.scale)
    const la = (e.latDeg * Math.PI) / 180
    const lo = (e.lonDeg * Math.PI) / 180
    extrusion = {
      color: e.color,
      dx: -d * Math.sin(lo),
      dy: d * Math.sin(la) * Math.cos(lo),
    }
  }

  return {
    lines,
    insets,
    anchor,
    fontScale,
    ...(lnSpcReduction ? { lnSpcReduction } : {}),
    contentHeight: result.contentHeight,
    ...(result.inkBottom ? { inkBottom: result.inkBottom } : {}),
    wrap,
    autofit: body.autofit ?? 'none',
    ...(extrusion ? { extrusion } : {}),
    ...(body.txWarp ? { txWarp: body.txWarp } : {}),
  }
}

/**
 * PowerPoint autofit's discrete shrink steps (fontScale in 7.5% increments +
 * lnSpcReduction in 0/10/20% bands), observed from real files
 * (92500/0, 85000/10000, 62500/20000…).
 * Step down and take the first that fits; floor 25%.
 */
const SHRINK_STEPS: Array<[number, number]> = [
  [0.925, 0],
  [0.85, 0.1],
  [0.775, 0.1],
  [0.7, 0.2],
  [0.625, 0.2],
  [0.55, 0.2],
  [0.475, 0.2],
  [0.4, 0.2],
  [0.325, 0.2],
  [0.25, 0.2],
]

// ── Rotated text (bodyPr vert="vert"/"vert270") ─────────────────────

/**
 * vert / vert270 ("rotate all text 90°/270°"): true whole-block rotation. The text is
 * laid out HORIZONTALLY in the 90°-swapped box (width↔height, insets swapped in
 * rotation-consistent pairs), then every glyph maps into real box coordinates with a
 * rotate90/rotate270 flag — the renderer rotates each glyph node, so CJK rotates with
 * the block (unlike eaVert's upright columns). Autofit/wrap/anchor run inside the
 * recursive horizontal pass and stay semantically correct; the anchor axis maps to the
 * horizontal stacking direction (top = right edge for vert, left edge for vert270,
 * matching PowerPoint). Line top/height stay in layout (pre-rotation) space — same
 * convention as eaVert columns, whose consumers skip them when `vert` is set.
 */
function layoutTextRotated(input: TextLayoutInput, vert: 'vert' | 'vert270'): RenderTextLayout {
  const { body, boxWidthPx, boxHeightPx, vp } = input
  const ins = {
    l: body.insets?.l ?? DEFAULT_INSETS_EMU.l,
    t: body.insets?.t ?? DEFAULT_INSETS_EMU.t,
    r: body.insets?.r ?? DEFAULT_INSETS_EMU.r,
    b: body.insets?.b ?? DEFAULT_INSETS_EMU.b,
  }
  const h = layoutText({
    ...input,
    body: {
      ...body,
      vert: undefined,
      // Swapped so the recursive avail dims equal the real box's cross dims:
      // layout width = boxH - t - b, layout height = boxW - l - r
      insets: { l: ins.t, r: ins.b, t: ins.l, b: ins.r },
    },
    boxWidthPx: boxHeightPx,
    boxHeightPx: boxWidthPx,
  })
  const realInsets = {
    l: emuToPx(ins.l, vp.scale),
    t: emuToPx(ins.t, vp.scale),
    r: emuToPx(ins.r, vp.scale),
    b: emuToPx(ins.b, vp.scale),
  }
  // Real content-area dims (glyph coords stay relative to the content-area top-left)
  const wc = Math.max(boxWidthPx - realInsets.l - realInsets.r, 1)
  const hc = Math.max(boxHeightPx - realInsets.t - realInsets.b, 1)
  // Plane map: vert (90° cw) sends layout (u,v) → (wc - v, u); vert270 (90° ccw) sends
  // (u,v) → (v, hc - u). The renderer positions a glyph node at (x, baselineY - 0.8em)
  // and rotates about that origin, so the emitted x/baselineY place the node origin at
  // the mapped glyph-box top-left (0.8em = Konva's browser-text baseline approximation,
  // the same constant glyphToDraw subtracts).
  const lines = h.lines.map((ln) => ({
    ...ln,
    runs: ln.runs.map((r) => {
      const topOff = 0.8 * r.fontSizePx
      return vert === 'vert'
        ? { ...r, x: wc - (r.baselineY - topOff), baselineY: r.x + topOff, rotate90: true }
        : { ...r, x: r.baselineY - topOff, baselineY: hc - r.x + topOff, rotate270: true }
    }),
  }))
  // contentHeight stays a REAL vertical extent (autofit-resize writes it into the shape
  // height): after rotation that is the pre-rotation layout's horizontal extent. The
  // layout-space inkBottom would likewise measure the wrong axis — drop it.
  let contentHeight = 0
  for (const ln of h.lines)
    for (const r of ln.runs) contentHeight = Math.max(contentHeight, r.x + r.widthPx)
  const { inkBottom: _layoutSpaceInk, ...rest } = h
  return { ...rest, lines, insets: realInsets, vert, contentHeight }
}

// ── Vertical text (bodyPr vert) ─────────────────────────────────────

/**
 * Vertical column layout (eaVert / wordArtVert): horizontal "lines" become "columns"
 * and characters within a column stack top to bottom. eaVert columns run right→left
 * and Latin words rotate 90° as whole words; wordArtVert ("stacked") columns run
 * left→right and every glyph stays upright, one letter on top of another. vert/vert270
 * rotate the whole block instead — see layoutTextRotated.
 * Approximations (v1):
 * - autofit doesn't self-iterate (only reuses the stored fontScale);
 *   justify/bidi/marL/super-subscript shifts are skipped.
 * Coordinate conventions match horizontal layout: GlyphRun.x/baselineY are relative
 * to the content area's (inside insets) top-left, zero renderer changes.
 * TextLine = one column; paraStart/srcRunIdx semantics unchanged, editor round-trips work.
 */
function layoutTextVertical(
  body: TextBody,
  vert: NonNullable<TextBody['vert']>,
  availWidth: number,
  availHeight: number,
  insets: { l: number; t: number; r: number; b: number },
  wrap: boolean,
  metrics: FontMetricsProvider,
  scale: number,
): RenderTextLayout {
  const fontScale = body.autofit === 'shrink' ? (body.fontScale ?? 1) : 1

  interface Col {
    runs: GlyphRun[] // baselineY fixed relative to column top; x parked at 0, backfilled by column position in pass two
    usedH: number
    colW: number
    paraStart: boolean
    align?: Paragraph['align']
    alignExplicit?: boolean
    rtl?: boolean
    softBreakAfter?: number
    /** Space-before/after maps to horizontal gaps between columns (carried by a paragraph's first/last column) */
    gapBefore: number
    gapAfter: number
  }
  const cols: Col[] = []
  let autoNum = 0

  for (const p of body.paragraphs) {
    const paraCols: Col[] = []
    let cur: GlyphRun[] = []
    let curH = 0
    let agg = { ascent: 0, descent: 0, size: 0 }

    const finishCol = (soft?: number) => {
      let size = agg.size
      if (!cur.length)
        size = ptToPx(
          fontScale !== 1 ? Math.max(1, Math.round(DEFAULT_SIZE_PT * fontScale)) : DEFAULT_SIZE_PT,
          scale,
        )
      paraCols.push({
        runs: cur,
        usedH: curH,
        // Line height maps to column width: 1.2em of the column's max font size × the paragraph line-spacing setting
        colW: lineH(p, size, scale, 0),
        paraStart: false,
        ...(soft != null ? { softBreakAfter: soft } : {}),
        gapBefore: 0,
        gapAfter: 0,
      })
      cur = []
      curH = 0
      agg = { ascent: 0, descent: 0, size: 0 }
    }

    const pushCell = (tok: Token, g: string, isBullet = false) => {
      const m = metrics.metrics(tok.style)
      // Vertical advance = the char style's line box height (CJK ≈ 1em; upright Latin also approximated by the line box)
      const adv = m.ascent + m.descent + tok.ls
      if (wrap && cur.length && curH + adv > availHeight) finishCol()
      agg = {
        ascent: Math.max(agg.ascent, m.ascent),
        descent: Math.max(agg.descent, m.descent),
        size: Math.max(agg.size, tok.style.fontSizePx),
      }
      cur.push({
        text: g,
        x: 0,
        baselineY: curH + m.ascent,
        fontFamily: metrics.displayFamily?.(tok.style, g) ?? tok.style.fontFamily,
        ...(tok.srcFont ? { srcFontFamily: tok.srcFont } : {}),
        fontSizePx: tok.style.fontSizePx,
        color: tok.color,
        bold: tok.style.bold,
        italic: tok.style.italic,
        underline: tok.underline,
        ...(tok.strike ? { strike: true } : {}),
        ...(tok.highlight ? { highlight: tok.highlight } : {}),
        ...(tok.outline ? { outline: tok.outline } : {}),
        ...(tok.gradient ? { gradient: tok.gradient } : {}),
        ...(tok.glow ? { glow: tok.glow } : {}),
        ...(tok.reflection ? { reflection: true } : {}),
        ...(tok.shadow ? { shadow: tok.shadow } : {}),
        widthPx: metrics.measure(g, tok.style),
        ...(tok.blPct ? { baselinePct: tok.blPct } : {}),
        ...(isBullet ? { isBullet: true } : { srcRunIdx: tok.srcRun }),
        ...(!isBullet && tok.link ? { link: tok.link } : {}),
        ascentPx: m.ascent,
      })
      curH += adv
    }

    // Latin words rotate 90° as whole words: vertical advance = word width (vertical length after rotation), column width contribution = line box height
    const pushRotated = (tok: Token) => {
      const m = metrics.metrics(tok.style)
      const adv = tokenWidth(tok, metrics)
      if (wrap && cur.length && curH + adv > availHeight) finishCol()
      agg = {
        ascent: Math.max(agg.ascent, m.ascent),
        descent: Math.max(agg.descent, m.descent),
        size: Math.max(agg.size, tok.style.fontSizePx),
      }
      cur.push({
        text: tok.text,
        x: 0,
        baselineY: curH + m.ascent,
        fontFamily: metrics.displayFamily?.(tok.style, tok.text) ?? tok.style.fontFamily,
        ...(tok.srcFont ? { srcFontFamily: tok.srcFont } : {}),
        fontSizePx: tok.style.fontSizePx,
        color: tok.color,
        bold: tok.style.bold,
        italic: tok.style.italic,
        underline: tok.underline,
        ...(tok.strike ? { strike: true } : {}),
        ...(tok.highlight ? { highlight: tok.highlight } : {}),
        ...(tok.outline ? { outline: tok.outline } : {}),
        ...(tok.gradient ? { gradient: tok.gradient } : {}),
        ...(tok.glow ? { glow: tok.glow } : {}),
        ...(tok.reflection ? { reflection: true } : {}),
        ...(tok.shadow ? { shadow: tok.shadow } : {}),
        widthPx: adv,
        // Rotated Latin words draw as whole strings too: keep draw kerning in step with the measure
        ...(tok.style.kerning === false ? { kerningOff: true } : {}),
        rotate90: true,
        srcRunIdx: tok.srcRun,
        ...(tok.link ? { link: tok.link } : {}),
        ascentPx: m.ascent,
      })
      curH += adv
    }

    const hasText = p.runs.some((r) => r.text.trim())
    const bulletType = p.bullet?.type
    const hasBullet = hasText && (bulletType === 'char' || bulletType === 'number')
    if (bulletType === 'number' && hasText)
      autoNum = autoNum === 0 ? (p.bullet?.startAt ?? 1) : autoNum + 1
    else if (bulletType !== 'number') autoNum = 0
    if (hasBullet && p.runs[0]) {
      const base = runStyle(p.runs[0], scale, fontScale)
      // <a:buSzPct>: bullet glyph size as a percentage of the first run's size
      let st =
        p.bullet?.sizePct != null
          ? { ...base, fontSizePx: base.fontSizePx * (p.bullet.sizePct / 100) }
          : base
      let glyph =
        bulletType === 'char' ? (p.bullet?.char ?? '•') : formatAutoNum(autoNum, p.bullet?.numType)
      const sym = bulletType === 'char' ? symbolBulletText(p.bullet?.font, glyph) : undefined
      if (sym) {
        glyph = sym
        st = { ...st, fontFamily: p.bullet!.font! }
      }
      pushCell(
        {
          text: '',
          style: st,
          color: p.bullet?.color ?? p.runs[0].color ?? '#000000',
          underline: false,
          ls: 0,
          breakable: false,
          isSpace: false,
          srcRun: 0,
        },
        glyph,
        true,
      )
    }

    for (const rawTok of tokenizeParagraph(p, scale, fontScale)) {
      const tok = substituteKerning(rawTok, metrics)
      if (tok.isBreak) {
        finishCol(tok.srcRun)
        continue
      }
      // Latin/digit words (no wide chars) rotate 90° whole per PowerPoint vertical-writing
      // semantics; CJK stays upright per char. wordArtVert ("stacked") keeps EVERY glyph
      // upright, one letter on top of another, so Latin also goes through the per-cell path.
      const hasWide = [...tok.text].some((ch) => isWideChar(ch.codePointAt(0) ?? 0))
      if (vert !== 'wordArtVert' && !hasWide && tok.text.trim()) {
        pushRotated(tok)
        continue
      }
      for (const g of graphemes(tok.text)) pushCell(tok, g)
    }
    if (cur.length || !paraCols.length) finishCol()

    paraCols[0]!.paraStart = true
    const singleW = paraCols[0]!.colW
    paraCols[0]!.gapBefore =
      ptToPx(p.spaceBefore ?? 0, scale) +
      (p.spaceBeforePct ? singleW * (p.spaceBeforePct / 100) : 0)
    paraCols[paraCols.length - 1]!.gapAfter =
      ptToPx(p.spaceAfter ?? 0, scale) + (p.spaceAfterPct ? singleW * (p.spaceAfterPct / 100) : 0)
    if (p.align) {
      for (const c of paraCols) {
        c.align = p.align
        c.alignExplicit = true
      }
    }
    if (paraBaseRtl(p)) for (const c of paraCols) c.rtl = true
    cols.push(...paraCols)
  }

  // anchor acts along the text flow: East Asian columns run right→left (top starts at
  // the right, bottom hugs the left); wordArtVert stacked columns run left→right per
  // ECMA-376 (wordArtVertRtl is the separate right-to-left variant)
  const ltr = vert === 'wordArtVert'
  const contentW = cols.reduce((a, c) => a + c.gapBefore + c.colW + c.gapAfter, 0)
  const anchor = body.anchor ?? 'top'
  const extraW = availWidth - contentW
  const anchorOff = anchor === 'middle' ? extraW / 2 : anchor === 'bottom' ? extraW : 0
  let xFlow = ltr ? anchorOff : availWidth - anchorOff
  let contentHeight = 0
  const lines: TextLine[] = cols.map((c) => {
    let colX: number
    if (ltr) {
      xFlow += c.gapBefore
      colX = xFlow
      xFlow = colX + c.colW + c.gapAfter
    } else {
      xFlow -= c.gapBefore
      colX = xFlow - c.colW
      xFlow = colX - c.gapAfter
    }
    // Paragraph alignment acts vertically within a column: ctr centers, r hugs the bottom (justify treated as left)
    const dy =
      c.align === 'center'
        ? (availHeight - c.usedH) / 2
        : c.align === 'right'
          ? availHeight - c.usedH
          : 0
    contentHeight = Math.max(contentHeight, dy + c.usedH)
    return {
      runs: c.runs.map((r) => ({
        ...r,
        // Rotated word anchor = column center shifted right by half the font size (after 90° clockwise the text box lands back on the column center)
        x: r.rotate90 ? colX + (c.colW + r.fontSizePx) / 2 : colX + (c.colW - r.widthPx) / 2,
        baselineY: r.baselineY + dy,
      })),
      top: dy,
      height: c.usedH,
      paraStart: c.paraStart,
      ...(c.softBreakAfter != null ? { softBreakAfter: c.softBreakAfter } : {}),
      ...(c.alignExplicit && c.align ? { align: c.align } : {}),
      ...(c.rtl ? { rtl: true } : {}),
    }
  })

  return { lines, insets, anchor, fontScale, contentHeight, wrap, vert }
}

/** Paragraph horizontal alignment → line x offset (center/right; justify treated as left for now). */
function alignOffset(align: Paragraph['align'], availWidth: number, lineWidth: number): number {
  if (align === 'center') return (availWidth - lineWidth) / 2
  if (align === 'right') return availWidth - lineWidth
  return 0
}

function layoutAll(
  body: TextBody,
  availWidth: number,
  wrap: boolean,
  metrics: FontMetricsProvider,
  scale: number,
  fontScale: number,
  lnSpcRed: number,
  trimEdgeSpacing?: boolean,
): { lines: TextLine[]; contentHeight: number; inkBottom?: number } {
  const outLines: TextLine[] = []
  // Glyph-extent bottom (last baseline + descent): PowerPoint sizes auto table rows by ink,
  // not by line-box sum (18pt probe: lnSpc 115% single-line row = 0.7333x box + descent)
  let inkBottom = 0
  let y = 0
  let autoNum = 0 // buAutoNum sequential numbering (reset on a non-numbered paragraph)
  for (const [pIdx, p] of body.paragraphs.entries()) {
    // Indent and bullets (body lines start at marL; the bullet
    // draws at marL+indent, negative indent = hanging indent; without a bullet the
    // first line starts at marL+indent). When no layer of the style chain defines
    // marL, PowerPoint's built-in default indents each level by 0.5" (457200 EMU).
    const marLPx = emuToPx(p.marL ?? (p.level ? p.level * 457200 : 0), scale)
    const indentPx = emuToPx(p.indent ?? 0, scale)
    const hasText = p.runs.some((r) => r.text.trim())
    const bulletType = p.bullet?.type
    const hasBullet = hasText && (bulletType === 'char' || bulletType === 'number')
    if (bulletType === 'number' && hasText)
      autoNum = autoNum === 0 ? (p.bullet?.startAt ?? 1) : autoNum + 1
    else if (bulletType !== 'number') autoNum = 0
    let bulletText =
      bulletType === 'char' ? (p.bullet?.char ?? '•') : formatAutoNum(autoNum, p.bullet?.numType)
    const symText = bulletType === 'char' ? symbolBulletText(p.bullet?.font, bulletText) : undefined
    if (symText) bulletText = symText

    const textX = marLPx
    const marRPx = emuToPx(p.marR ?? 0, scale)
    const avail = Math.max(availWidth - textX - marRPx, 1)
    // RTL base direction mirrors the paragraph box (PowerPoint semantics, probe-measured):
    // marL/indent measure from the RIGHT edge, the bullet glyph hangs on the right, and the
    // default alignment is right. Explicit algn values stay physical (l = left edge).
    const mirror = paraBaseRtl(p)
    // Affects layout only, not written back to TextLine.align (an editor commit would
    // store it as an explicit value)
    const align = p.align ?? (mirror ? ('right' as const) : undefined)
    // Bullet glyph style/advance (drawn on the first line only). PowerPoint reserves
    // the glyph's advance like a tab stop: body text can never start before the glyph
    // ends, even when the glyph is wider than the hanging indent (-indent).
    let bulletSt: RunStyle | undefined
    let bulletW = 0
    if (hasBullet) {
      const base = runStyle(p.runs[0]!, scale, fontScale)
      // <a:buSzPct>: bullet glyph size as a percentage of the first run's size
      bulletSt =
        p.bullet?.sizePct != null
          ? { ...base, fontSizePx: base.fontSizePx * (p.bullet.sizePct / 100) }
          : base
      if (symText) bulletSt = { ...bulletSt, fontFamily: p.bullet!.font! }
      bulletW = metrics.measure(bulletText, bulletSt)
    }
    const bulletX = Math.max(marLPx + indentPx, 0)
    // Glyph wider than the hanging indent: the first line's text start shifts right by
    // this much, so the first line must also wrap that much earlier
    const bulletOverflowPx = hasBullet ? Math.max(bulletX + bulletW - textX, 0) : 0
    // The first line's x shift: bullet-overflow push, or the first-line indent itself —
    // it consumes (negative: adds) that much of the first line's wrap budget
    const firstLineDx = hasBullet ? bulletOverflowPx : indentPx
    const laid = layoutParagraph(p, avail, wrap, metrics, scale, fontScale, lnSpcRed, firstLineDx, {
      stopsPx: (p.tabStops ?? []).map((t) => emuToPx(t.pos, scale)),
      defaultPx: Math.max(emuToPx(p.defTabSz ?? 914400, scale), 1),
      originPx: textX,
    })
    // Space before/after: spcPts is absolute pt; spcPct is a percentage of the paragraph's single line height (100 = one line).
    // PowerPoint ignores space-before on a text frame's FIRST paragraph in every body
    // (0047 measured: defaultTextStyle spcBef 50% shifted no first line), not just table cells.
    const singleH = laid[0]?.singleH ?? 0
    if (pIdx !== 0)
      y +=
        ptToPx(p.spaceBefore ?? 0, scale) +
        (p.spaceBeforePct ? singleH * (p.spaceBeforePct / 100) : 0)
    laid.forEach((ln, li) => {
      const baseline = y + (ln.leadAbove ?? 0) + ln.ascent
      inkBottom = Math.max(inkBottom, baseline + ln.descent)
      const lineWidth = ln.runs.reduce((acc, r) => acc + r.widthPx, 0)
      // Without a bullet the first line adds indent (positive or negative); with a bullet
      // the body starts at marL, pushed right when the glyph overflows the hanging indent
      const firstShift = !hasBullet && li === 0 ? indentPx : 0
      const bulletShift = li === 0 ? bulletOverflowPx : 0
      // justify: lines filled by wrapping (not paragraph-final, not hard breaks) spread
      // the remaining width into word gaps (U+0020 only), like PowerPoint — letter
      // spacing stays untouched. Lines with no space (CJK) spread between characters
      // instead, via justifyExtraPx (a draw-only field) so letterSpacing round-trips
      // stay clean. Bidi lines keep the character spread: fragment repositioning
      // assumes LTR visual order.
      let lineRuns = ln.runs
      if (
        align === 'justify' &&
        wrap &&
        li < laid.length - 1 &&
        ln.softBreakAfter == null &&
        ln.runs.length
      ) {
        const extra = avail - firstShift - bulletShift - lineWidth
        // line-trailing spaces get no share (nothing follows them)
        const last = ln.runs.length - 1
        const spaceCount = ln.runs.reduce(
          (acc, r, i) =>
            acc + ((i === last ? r.text.replace(/ +$/, '') : r.text).match(/ /g)?.length ?? 0),
          0,
        )
        if (extra > 0 && spaceCount > 0 && !ln.runs.some((r) => r.rtl)) {
          const per = extra / spaceCount
          let remaining = spaceCount
          let shift = 0
          lineRuns = ln.runs.flatMap((r) => {
            if (remaining <= 0 || !r.text.includes(' ')) return [{ ...r, x: r.x + shift }]
            const style = {
              fontFamily: r.fontFamily,
              fontSizePx: r.fontSizePx,
              bold: r.bold,
              italic: r.italic,
            }
            const ls = r.letterSpacingPx ?? 0
            const frags = r.text.match(/[^ ]+ *| +/g) ?? [r.text]
            const widths = frags.map((f) => metrics.measure(f, style) + ls * [...f].length)
            const wSum = widths.reduce((a, b) => a + b, 0)
            // normalize so fragment widths keep summing to the measured run width
            const norm = wSum > 0 ? r.widthPx / wSum : 1
            let fx = r.x + shift
            return frags.map((f, i) => {
              const natural = widths[i]! * norm
              const nSp = Math.min(/ +$/.exec(f)?.[0].length ?? 0, remaining)
              remaining -= nSp
              const frag = { ...r, text: f, x: fx, widthPx: natural + nSp * per }
              fx += natural + nSp * per
              shift += nSp * per
              return frag
            })
          })
        } else {
          const totalChars = ln.runs.reduce((acc, r) => acc + [...r.text].length, 0)
          if (totalChars > 1 && extra > 0) {
            const per = extra / (totalChars - 1)
            let consumed = 0
            lineRuns = ln.runs.map((r) => {
              const chars = [...r.text].length
              const jr = {
                ...r,
                x: r.x + consumed * per,
                widthPx: r.widthPx + chars * per,
                justifyExtraPx: per,
              }
              consumed += chars
              return jr
            })
          }
        }
      }
      // Center/right alignment counts the overflow push so the text stays inside the box
      const off = alignOffset(align, avail, lineWidth + bulletShift)
      let dx = textX + firstShift + bulletShift + off
      // Mirrored box: the body's right boundary sits marL (+ first-line shifts) from the
      // right edge; every line lives inside [0, rightEdge]. Spread justified lines fill
      // that whole span (the spread width already equals it), unspread ones (paragraph-
      // final, hard breaks) go flush right like PowerPoint.
      const rightEdge = availWidth - textX - firstShift - bulletShift
      const mirrorSpread = mirror && align === 'justify' && lineRuns !== ln.runs
      if (mirror) {
        // Mirrored: marL measures from the right edge, marR from the left
        dx =
          mirrorSpread || align === 'left'
            ? marRPx
            : align === 'center'
              ? marRPx + (rightEdge - marRPx - lineWidth) / 2
              : rightEdge - lineWidth
      }
      const runs = lineRuns.map((r) => ({
        ...r,
        x: r.x + dx,
        baselineY: baseline - (r.baselineShiftPx ?? 0),
      }))
      if (hasBullet && li === 0) {
        const st = bulletSt!
        // Mirrored: the glyph keeps its reserved advance but on the right of the body text
        // (right edge at availWidth - bulletX when the line is flush right); numbered
        // glyphs render with an RTL base so "1." displays as ".1" like PowerPoint
        const bx = mirror
          ? dx + (mirrorSpread ? rightEdge : lineWidth) + (textX + bulletShift - bulletX - bulletW)
          : bulletX + off
        runs.unshift({
          text: bulletText,
          x: bx,
          ...(mirror ? { rtl: true } : {}),
          baselineY: baseline,
          fontFamily: metrics.displayFamily?.(st, bulletText) ?? st.fontFamily,
          fontSizePx: st.fontSizePx,
          color: p.bullet?.color ?? p.runs[0]?.color ?? '#000000',
          bold: st.bold,
          italic: false,
          underline: false,
          widthPx: bulletW,
          isBullet: true,
          ascentPx: metrics.metrics(st).ascent,
        })
      }
      outLines.push({
        runs,
        top: y,
        height: ln.height,
        ...(ln.leadAbove ? { leadAbove: ln.leadAbove } : {}),
        paraStart: li === 0,
        ...(ln.trailingSpace ? { trailingSpace: true } : {}),
        ...(ln.trailingText ? { trailingText: ln.trailingText } : {}),
        ...(ln.softBreakAfter != null ? { softBreakAfter: ln.softBreakAfter } : {}),
        ...(p.align ? { align: p.align } : {}),
        ...(paraBaseRtl(p) ? { rtl: true } : {}),
        ...(p.level ? { level: p.level } : {}),
        ...(marLPx ? { marLPx } : {}),
        ...(indentPx ? { indentPx } : {}),
      })
      y += ln.height
    })
    if (!(trimEdgeSpacing && pIdx === body.paragraphs.length - 1))
      y +=
        ptToPx(p.spaceAfter ?? 0, scale) + (p.spaceAfterPct ? singleH * (p.spaceAfterPct / 100) : 0)
  }
  return { lines: outLines, contentHeight: y, inkBottom }
}
