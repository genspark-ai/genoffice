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
  return {
    fontFamily: run.fontFamily || DEFAULT_FONT,
    fontSizePx: ptToPx(sizePt, scale) * fontScale,
    bold: !!run.bold,
    italic: !!run.italic,
  }
}

/** Token width = font advance width + letter spacing × char count (matches canvas letterSpacing: appended after each char) */
function tokenWidth(tok: Token, metrics: FontMetricsProvider): number {
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
      } else if (ch === ' ' || ch === '\t') {
        flushWord()
        tokens.push({ ...base, text: ch, breakable: true, isSpace: true })
      } else if (ch === ' ') {
        // NBSP: glues neighbors into one token but draws/measures at plain-space
        // width (fonts like Carlito have no U+00A0 glyph → the missing-glyph
        // fallback would badly over-measure it)
        buf += ' '
      } else if (isWideChar(cp)) {
        flushWord()
        tokens.push({ ...base, text: ch, breakable: true, isSpace: false })
      } else {
        buf += ch
      }
    }
    flushWord()
  })
  return tokens
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
function applyBidi(tokens: Token[]): Token[] {
  const text = tokens.map((t) => t.text).join('')
  if (!RTL_RE.test(text)) return tokens
  bidiApi ??= bidiFactory()
  const { levels } = bidiApi.getEmbeddingLevels(text)
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

/** Paragraph base direction: decided by the first strong-direction character (used for the default alignment; RTL paragraphs without explicit alignment align right). */
function paraBaseRtl(p: Paragraph): boolean {
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
  /** First-line width reduction (px): a bullet glyph overflowing the hanging indent
   *  pushes the first line's text start right, so it must wrap that much earlier. */
  firstLineShrinkPx = 0,
): LaidLine[] {
  const tokens = applyBidi(tokenizeParagraph(p, scale, fontScale)).map((tok, logicalOrder) => ({
    ...tok,
    logicalOrder,
  }))
  const lines: LaidLine[] = []
  let cur: Token[] = []
  let curW = 0

  const pushLine = (toks: Token[]) => {
    // Strip trailing whitespace (recorded: the editor re-adds a space when joining wrapped lines back into a paragraph)
    let trailingSpace = false
    while (toks.length && toks[toks.length - 1]!.isSpace) {
      toks.pop()
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
            fontSizePx: ptToPx(DEFAULT_SIZE_PT, scale) * fontScale,
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
        ...(trailingSpace ? { trailingSpace } : {}),
      })
      return
    }
    const line = buildLine(toks, metrics, p, scale, lnSpcRed)
    if (trailingSpace) line.trailingSpace = true
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
    const w = tokenWidth(tok, metrics)
    // The first line loses firstLineShrinkPx to the overflowing bullet glyph; evaluated
    // lazily because the soft wrap right below can end line 0 for this same token
    const lineAvail = () => (lines.length === 0 ? availWidth - firstLineShrinkPx : availWidth)
    if (wrap && cur.length && curW + w > lineAvail() && !tok.isSpace) {
      pushLine(cur)
      cur = []
      curW = 0
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
      ...(tok.outline ? { outline: tok.outline } : {}),
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
    l: emuToPx(body.insets?.l ?? 91440, vp.scale),
    t: emuToPx(body.insets?.t ?? 45720, vp.scale),
    r: emuToPx(body.insets?.r ?? 91440, vp.scale),
    b: emuToPx(body.insets?.b ?? 45720, vp.scale),
  }
  const availWidth = Math.max(boxWidthPx - insets.l - insets.r, 1)
  const availHeight = Math.max(boxHeightPx - insets.t - insets.b, 1)
  const wrap = body.wrap !== false

  if (body.vert)
    return layoutTextVertical(
      body,
      body.vert,
      availWidth,
      availHeight,
      insets,
      wrap,
      metrics,
      vp.scale,
    )

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
  if (body.autofit === 'shrink' && fitSpan(result) > fitTarget * 1.03) {
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
  const lines = dy
    ? result.lines.map((ln) => ({
        ...ln,
        top: ln.top + dy,
        runs: ln.runs.map((r) => ({ ...r, baselineY: r.baselineY + dy })),
      }))
    : result.lines

  return {
    lines,
    insets,
    anchor,
    fontScale,
    ...(lnSpcReduction ? { lnSpcReduction } : {}),
    contentHeight: result.contentHeight,
    ...(result.inkBottom ? { inkBottom: result.inkBottom } : {}),
    wrap,
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

// ── Vertical text (bodyPr vert) ─────────────────────────────────────

/**
 * Vertical column layout: horizontal "lines" become "columns", columns run right
 * to left, and characters within a column stack top to bottom.
 * Approximations (v1):
 * - vert/vert270/wordArtVert should rotate the whole box 90°/270°, but that is too
 *   invasive for rendering/hit-testing/editing overlays, so degrade to the same
 *   column layout as eaVert;
 * - PowerPoint rotates Latin characters 90° under eaVert; here they stay upright
 *   (one grapheme per cell);
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
      if (!cur.length) size = ptToPx(DEFAULT_SIZE_PT, scale) * fontScale
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
        ...(tok.shadow ? { shadow: tok.shadow } : {}),
        widthPx: adv,
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
      let glyph = bulletType === 'char' ? (p.bullet?.char ?? '•') : `${autoNum}.`
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

    for (const tok of tokenizeParagraph(p, scale, fontScale)) {
      if (tok.isBreak) {
        finishCol(tok.srcRun)
        continue
      }
      // Latin/digit words (no wide chars) rotate 90° whole per PowerPoint vertical-writing semantics; CJK stays upright per char
      const hasWide = [...tok.text].some((ch) => isWideChar(ch.codePointAt(0) ?? 0))
      if (!hasWide && tok.text.trim()) {
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
    cols.push(...paraCols)
  }

  // anchor acts along the text flow (right→left): top starts at the right, middle centers, bottom hugs the left
  const contentW = cols.reduce((a, c) => a + c.gapBefore + c.colW + c.gapAfter, 0)
  const anchor = body.anchor ?? 'top'
  const extraW = availWidth - contentW
  let xRight = availWidth - (anchor === 'middle' ? extraW / 2 : anchor === 'bottom' ? extraW : 0)
  let contentHeight = 0
  const lines: TextLine[] = cols.map((c) => {
    xRight -= c.gapBefore
    const colX = xRight - c.colW
    xRight = colX - c.gapAfter
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
    let bulletText = bulletType === 'char' ? (p.bullet?.char ?? '•') : `${autoNum}.`
    const symText = bulletType === 'char' ? symbolBulletText(p.bullet?.font, bulletText) : undefined
    if (symText) bulletText = symText

    const textX = marLPx
    const avail = Math.max(availWidth - textX, 1)
    // Paragraphs with an RTL base direction (first strong char is RTL) default to right
    // alignment when none is explicit; affects layout only, not written back to
    // TextLine.align (an editor commit would store it as an explicit value)
    const align = p.align ?? (paraBaseRtl(p) ? ('right' as const) : undefined)
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
    const laid = layoutParagraph(
      p,
      avail,
      wrap,
      metrics,
      scale,
      fontScale,
      lnSpcRed,
      bulletOverflowPx,
    )
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
      // the remaining width into character spacing; the spread goes through
      // justifyExtraPx (a draw-only field) so letterSpacing round-trips stay clean
      let lineRuns = ln.runs
      if (
        align === 'justify' &&
        wrap &&
        li < laid.length - 1 &&
        ln.softBreakAfter == null &&
        ln.runs.length
      ) {
        const totalChars = ln.runs.reduce((acc, r) => acc + [...r.text].length, 0)
        const extra = avail - firstShift - bulletShift - lineWidth
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
      // Center/right alignment counts the overflow push so the text stays inside the box
      const off = alignOffset(align, avail, lineWidth + bulletShift)
      const dx = textX + firstShift + bulletShift + off
      const runs = lineRuns.map((r) => ({
        ...r,
        x: r.x + dx,
        baselineY: baseline - (r.baselineShiftPx ?? 0),
      }))
      if (hasBullet && li === 0) {
        const st = bulletSt!
        runs.unshift({
          text: bulletText,
          x: bulletX + off,
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
        ...(ln.softBreakAfter != null ? { softBreakAfter: ln.softBreakAfter } : {}),
        ...(p.align ? { align: p.align } : {}),
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
