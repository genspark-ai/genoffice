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
  /** Run hyperlink (TextRun.hyperlink encoding) */
  link?: string
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
    // Per grapheme cluster: combining marks / ZWJ emoji always share a token with their base char, never split by wrapping
    for (const ch of graphemes(run.text)) {
      const cp = ch.codePointAt(0) ?? 0
      if (ch === '\n' || ch === '\v') {
        flushWord()
        tokens.push({ ...base, text: '\n', breakable: true, isSpace: false, isBreak: true })
      } else if (ch === ' ' || ch === '\t') {
        flushWord()
        tokens.push({ ...base, text: ch, breakable: true, isSpace: true })
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
        height: lineH(p, m, scale, lnSpcRed),
        ascent: m.ascent,
        descent: m.descent,
        singleH: Math.max(m.lineHeight, m.ascent + m.descent),
        ...(trailingSpace ? { trailingSpace } : {}),
      })
      return
    }
    const line = buildLine(toks, metrics, p, scale, lnSpcRed)
    if (trailingSpace) line.trailingSpace = true
    lines.push(line)
  }

  for (const tok of tokens) {
    // <a:br/> forced break: breaks regardless of wrap; record the sentinel run index for editor round-trips
    if (tok.isBreak) {
      pushLine(cur)
      lines[lines.length - 1]!.softBreakAfter = tok.srcRun
      cur = []
      curW = 0
      continue
    }
    const w = tokenWidth(tok, metrics)
    if (wrap && cur.length && curW + w > availWidth && !tok.isSpace) {
      pushLine(cur)
      cur = []
      curW = 0
    }
    // Hard-break over-long words (a single token wider than the line)
    if (wrap && !cur.length && w > availWidth && tok.text.length > 1 && !tok.isSpace) {
      for (const seg of hardBreak(tok, availWidth, metrics)) {
        pushLine([seg])
      }
      continue
    }
    cur.push(tok)
    curW += w
  }
  if (cur.length) pushLine(cur)
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
  let lineHM = 0
  const runs: GlyphRun[] = []
  // Lines with RTL levels are reordered into visual order before laying out x (break in logical order → draw in visual order)
  for (const tok of visualOrder(toks)) {
    const m = metrics.metrics(tok.style)
    ascent = Math.max(ascent, m.ascent)
    descent = Math.max(descent, m.descent)
    lineHM = Math.max(lineHM, m.lineHeight)
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
      widthPx: w,
      ...(tok.ls ? { letterSpacingPx: tok.ls } : {}),
      ...(tok.outline ? { outline: tok.outline } : {}),
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
  const height = lineH(
    p,
    { ascent, descent, lineHeight: lineHM || ascent + descent },
    scale,
    lnSpcRed,
  )
  return {
    runs,
    height,
    ascent,
    descent,
    singleH: Math.max(lineHM || ascent + descent, ascent + descent),
  }
}

/**
 * Computes line height with the paragraph's line-spacing settings. Single spacing =
 * the font's suggested line height (opentype: hhea ascent+descent; heuristic: 1.2em)
 * — with no extra ×1.2 on top (the earlier double scaling made large-size
 * title spacing ~20% too tall).
 */
function lineH(
  p: Paragraph,
  m: { ascent: number; descent: number; lineHeight: number },
  scale: number,
  lnSpcRed = 0,
): number {
  const single = Math.max(m.lineHeight, m.ascent + m.descent)
  if (p.lineExact != null) return ptToPx(p.lineExact, scale) // exact line spacing is unaffected by autofit spacing reduction
  const base = p.lineHeight != null ? single * (p.lineHeight / 100) : single
  return base * (1 - lnSpcRed)
}

export interface TextLayoutInput {
  body: TextBody
  boxWidthPx: number
  boxHeightPx: number
  metrics: FontMetricsProvider
  vp: Viewport
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

  const build = (fontScale: number, lnSpcRed: number) =>
    layoutAll(body, availWidth, wrap, metrics, vp.scale, fontScale, lnSpcRed)

  // PowerPoint's stored shrink ratio takes priority (files with shrunk text render
  // as-is; we don't scale back up per our own metrics). Only if content still
  // overflows (we edited text / metric differences) do we step further down the
  // discrete autofit ladder (see SHRINK_STEPS), capped at the stored value
  const storedScale = body.autofit === 'shrink' ? (body.fontScale ?? 1) : 1
  const storedRed = body.autofit === 'shrink' ? (body.lnSpcReduction ?? 0) : 0
  let fontScale = storedScale
  let lnSpcReduction = storedRed
  let result = build(fontScale, lnSpcReduction)
  if (body.autofit === 'shrink' && result.contentHeight > availHeight) {
    for (const [fs, red] of SHRINK_STEPS) {
      if (fs >= storedScale - 1e-6) continue
      const effRed = Math.max(red, storedRed)
      const r = build(fs, effRed)
      fontScale = fs
      lnSpcReduction = effRed
      result = r
      if (r.contentHeight <= availHeight) break
      // Still doesn't fit after all steps → stop at 25% (overflow allowed)
    }
  }

  // Vertical anchor: middle/bottom shift everything down (negative offsets allowed
  // when content is too tall → overflow upward/on both sides),
  // baked straight into line coordinates so the renderer draws as-is.
  const anchor = body.anchor ?? 'top'
  const extraH = availHeight - result.contentHeight
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
    let agg = { ascent: 0, descent: 0, lineHeight: 0 }

    const finishCol = (soft?: number) => {
      let m = agg
      if (!cur.length) {
        const st: RunStyle = {
          fontFamily: DEFAULT_FONT,
          fontSizePx: ptToPx(DEFAULT_SIZE_PT, scale) * fontScale,
          bold: false,
          italic: false,
        }
        m = metrics.metrics(st)
      }
      paraCols.push({
        runs: cur,
        usedH: curH,
        // Line height maps to column width: the column's max metrics × the paragraph line-spacing setting
        colW: lineH(p, { ...m, lineHeight: m.lineHeight || m.ascent + m.descent }, scale, 0),
        paraStart: false,
        ...(soft != null ? { softBreakAfter: soft } : {}),
        gapBefore: 0,
        gapAfter: 0,
      })
      cur = []
      curH = 0
      agg = { ascent: 0, descent: 0, lineHeight: 0 }
    }

    const pushCell = (tok: Token, g: string, isBullet = false) => {
      const m = metrics.metrics(tok.style)
      // Vertical advance = the char style's line box height (CJK ≈ 1em; upright Latin also approximated by the line box)
      const adv = m.ascent + m.descent + tok.ls
      if (wrap && cur.length && curH + adv > availHeight) finishCol()
      agg = {
        ascent: Math.max(agg.ascent, m.ascent),
        descent: Math.max(agg.descent, m.descent),
        lineHeight: Math.max(agg.lineHeight, m.lineHeight),
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
        ...(tok.outline ? { outline: tok.outline } : {}),
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
        lineHeight: Math.max(agg.lineHeight, m.lineHeight),
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
        ...(tok.outline ? { outline: tok.outline } : {}),
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
    if (bulletType === 'number' && hasText) autoNum += 1
    else if (bulletType !== 'number') autoNum = 0
    if (hasBullet && p.runs[0]) {
      const base = runStyle(p.runs[0], scale, fontScale)
      // <a:buSzPct>: bullet glyph size as a percentage of the first run's size
      const st =
        p.bullet?.sizePct != null
          ? { ...base, fontSizePx: base.fontSizePx * (p.bullet.sizePct / 100) }
          : base
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
        bulletType === 'char' ? (p.bullet?.char ?? '•') : `${autoNum}.`,
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
): { lines: TextLine[]; contentHeight: number } {
  const outLines: TextLine[] = []
  let y = 0
  let autoNum = 0 // buAutoNum sequential numbering (reset on a non-numbered paragraph)
  for (const p of body.paragraphs) {
    // Indent and bullets (body lines start at marL; the bullet
    // draws at marL+indent, negative indent = hanging indent; without a bullet the
    // first line starts at marL+indent)
    const marLPx = emuToPx(p.marL ?? 0, scale)
    const indentPx = emuToPx(p.indent ?? 0, scale)
    const hasText = p.runs.some((r) => r.text.trim())
    const bulletType = p.bullet?.type
    const hasBullet = hasText && (bulletType === 'char' || bulletType === 'number')
    if (bulletType === 'number' && hasText) autoNum += 1
    else if (bulletType !== 'number') autoNum = 0
    const bulletText = bulletType === 'char' ? (p.bullet?.char ?? '•') : `${autoNum}.`

    const textX = marLPx
    const avail = Math.max(availWidth - textX, 1)
    // Paragraphs with an RTL base direction (first strong char is RTL) default to right
    // alignment when none is explicit; affects layout only, not written back to
    // TextLine.align (an editor commit would store it as an explicit value)
    const align = p.align ?? (paraBaseRtl(p) ? ('right' as const) : undefined)
    const laid = layoutParagraph(p, avail, wrap, metrics, scale, fontScale, lnSpcRed)
    // Space before/after: spcPts is absolute pt; spcPct is a percentage of the paragraph's single line height (100 = one line)
    const singleH = laid[0]?.singleH ?? 0
    y +=
      ptToPx(p.spaceBefore ?? 0, scale) +
      (p.spaceBeforePct ? singleH * (p.spaceBeforePct / 100) : 0)
    laid.forEach((ln, li) => {
      const baseline = y + ln.ascent
      const lineWidth = ln.runs.reduce((acc, r) => acc + r.widthPx, 0)
      // Without a bullet the first line adds indent (positive or negative); with a bullet the body always starts at marL
      const firstShift = !hasBullet && li === 0 ? indentPx : 0
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
        const extra = avail - firstShift - lineWidth
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
      const dx = textX + firstShift + alignOffset(align, avail, lineWidth)
      const runs = lineRuns.map((r) => ({
        ...r,
        x: r.x + dx,
        baselineY: baseline - (r.baselineShiftPx ?? 0),
      }))
      if (hasBullet && li === 0) {
        const base = runStyle(p.runs[0]!, scale, fontScale)
        // <a:buSzPct>: bullet glyph size as a percentage of the first run's size
        const st =
          p.bullet?.sizePct != null
            ? { ...base, fontSizePx: base.fontSizePx * (p.bullet.sizePct / 100) }
            : base
        const bw = metrics.measure(bulletText, st)
        const bx = Math.max(marLPx + indentPx, 0) + alignOffset(align, avail, lineWidth)
        runs.unshift({
          text: bulletText,
          x: bx,
          baselineY: baseline,
          fontFamily: metrics.displayFamily?.(st, bulletText) ?? st.fontFamily,
          fontSizePx: st.fontSizePx,
          color: p.bullet?.color ?? p.runs[0]?.color ?? '#000000',
          bold: st.bold,
          italic: false,
          underline: false,
          widthPx: bw,
          isBullet: true,
          ascentPx: metrics.metrics(st).ascent,
        })
      }
      outLines.push({
        runs,
        top: y,
        height: ln.height,
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
    y +=
      ptToPx(p.spaceAfter ?? 0, scale) + (p.spaceAfterPct ? singleH * (p.spaceAfterPct / 100) : 0)
  }
  return { lines: outLines, contentHeight: y }
}
