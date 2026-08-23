/**
 * Line-width preflight (P21 A): rebuilt paragraphs render in the OUTPUT font,
 * whose advances can run wider than the source's embedded face even after the
 * metric-compatible mapping (fontmap.ts) — and the rebuilt column itself can
 * be a hair narrower than the source's (margins clamp, column widths scale
 * into the content width). A line that no longer fits wraps one word early,
 * every extra wrap compounds down the page, and exact line heights turn the
 * surplus into spilled pages.
 *
 * Each block's lines are re-measured with the output font's real advances
 * against the width the paragraph will ACTUALLY get (its output column /
 * cell width — the caller passes it, only the rebuild loop knows). When a
 * line overflows, the whole block tightens uniformly — first negative
 * w:spacing (the P14 channel) up to a visual cap, then 0.5pt font-size steps
 * down to a floor near the source size (the P8 calibration already pinned
 * that size to the rendered ink).
 *
 * Coverage (P31): Latin measures through the resolved output face, CJK
 * through synthetic 1-em fullwidth advances (font-independent), Arabic
 * through isolated advances scaled by a calibrated shaping ratio. Blocks in
 * other complex scripts stay out — no shaping model for them.
 */
import { advanceWidths } from '../../../font-metrics/src/index'
import type { Span, TextBlock } from '../ir'

/** twips per point (w:spacing emission grid) */
const PT_TO_TWIPS = 20

/** measured line width must exceed avail by both gates before anything moves —
 * un-kerned advances overestimate slightly (LO kerns and ligates by default) */
const OVERFLOW_RATIO = 1.0
const OVERFLOW_ABS_PT = 0.25
/** tightening caps: spacing beyond ~0.05em reads as squeezed text, and the
 * artifact detector treats ≤ -0.15em source tracking as a metric lie (P14) */
const SPACING_CAP_EMS = 0.05
/** font shrink: 0.5pt steps, at most 3, never below 6pt or 85% of the source */
const SHRINK_STEP_PT = 0.5
const SHRINK_MAX_STEPS = 3
const SHRINK_FLOOR_PT = 6
const SHRINK_FLOOR_RATIO = 0.85

/** pt width of `text` in the output font, or null when unmeasurable */
export type SpanMeasurer = (
  text: string,
  family: string,
  sizePt: number,
  bold: boolean,
  italic: boolean,
) => number | null

export const systemFontMeasurer: SpanMeasurer = (text, family, sizePt, bold, italic) => {
  const widths = advanceWidths(family, text, sizePt, { bold, italic })
  if (!widths) return null
  const total = widths.reduce((a, b) => a + b, 0)
  return Number.isNaN(total) ? null : total / 20
}

// ── CJK synthetic advances (P31 A) ──
// Every CJK output face renders ideographs, kana and hangul at exactly 1 em
// and fullwidth punctuation likewise, so their width needs no font file at
// all — which also sidesteps unresolvable mapped families. ASCII inside a
// CJK-script span (rare) approximates at the half-width slot.

/** fullwidth codepoint test: CJK ideographs, kana, hangul, fullwidth forms */
const isFullwidthCode = (code: number): boolean =>
  code === 0x2015 || // CJK dash (P31 D normalization target), 1 em in every EA face
  (code >= 0x1100 && code <= 0x115f) || // hangul jamo
  (code >= 0x2e80 && code <= 0x9fff) || // radicals, kana, CJK ideographs
  (code >= 0xa000 && code <= 0xa4cf) || // yi
  (code >= 0xac00 && code <= 0xd7a3) || // hangul syllables
  (code >= 0xf900 && code <= 0xfaff) || // compat ideographs
  (code >= 0xfe30 && code <= 0xfe4f) || // CJK compat forms
  (code >= 0xff00 && code <= 0xff60) || // fullwidth forms
  (code >= 0xffe0 && code <= 0xffe6) ||
  (code >= 0x20000 && code <= 0x3fffd) // ext ideographs

const cjkSyntheticWidthPt = (text: string, sizePt: number): number => {
  let ems = 0
  for (const ch of text) ems += isFullwidthCode(ch.codePointAt(0)!) ? 1 : 0.5
  return ems * sizePt
}

// ── Arabic shaped-width estimate (P31 B) ──
// advanceWidths sums ISOLATED-form advances; Word shapes the run, and the
// joined initial/medial forms run narrower. Calibrated via CoreText shaped
// line bounds vs isolated sums on Arial (0.72–0.78 across textbook phrases);
// erring LOW under-tightens, which is the safe direction for an estimate.
const ARABIC_SHAPED_RATIO = 0.75
/** measurement stand-ins when the mapped family cannot render Arabic */
const ARABIC_FALLBACK_FAMILIES = ['Geeza Pro', 'Arial', 'Times New Roman']

const CJK_SCRIPTS = new Set(['cjk', 'kana', 'hangul'])
const MEASURABLE_SCRIPTS = new Set(['latin', 'common', 'arabic', ...CJK_SCRIPTS])

// a span without a family renders in the document default (Calibri); the
// measurable stand-in is Arial, whose advances run ~5% wider (P31 C)
const DEFAULT_LATIN_STANDIN = 'Arial'
const CALIBRI_VS_ARIAL = 0.95

/** measured width of one span in the output face, script-dispatched */
const spanWidthPt = (
  s: Span,
  text: string,
  sizePt: number,
  measure: SpanMeasurer,
): number | null => {
  if (CJK_SCRIPTS.has(s.script)) return cjkSyntheticWidthPt(text, sizePt)
  if (s.script === 'arabic') {
    let w = s.fontFamily === '' ? null : measure(text, s.fontFamily, sizePt, s.bold, s.italic)
    for (const fam of ARABIC_FALLBACK_FAMILIES) {
      if (w !== null) break
      w = measure(text, fam, sizePt, s.bold, s.italic)
    }
    return w === null ? null : w * ARABIC_SHAPED_RATIO
  }
  if (s.fontFamily === '') {
    const w = measure(text, DEFAULT_LATIN_STANDIN, sizePt, s.bold, s.italic)
    return w === null ? null : w * CALIBRI_VS_ARIAL
  }
  return measure(text, s.fontFamily, sizePt, s.bold, s.italic)
}

const measurable = (block: TextBlock): boolean =>
  (block.dir === 'ltr' || block.dir === 'rtl') &&
  block.lines.length >= 1 &&
  block.tocEntry === undefined &&
  block.list === undefined &&
  block.cardId === undefined &&
  block.lines.every((line) =>
    line.spans.every(
      (s) =>
        s.noteRef !== undefined ||
        s.invisible === true ||
        (MEASURABLE_SCRIPTS.has(s.script) && !s.text.includes('\t')),
    ),
  )

interface MeasuredLine {
  widthPt: number
  codepoints: number
}

/** trailing spaces hang past the wrap edge in Word — they never force a wrap */
const measureLine = (
  spans: readonly Span[],
  measure: SpanMeasurer,
  sizeScale: number,
): MeasuredLine | null => {
  let widthPt = 0
  let codepoints = 0
  for (const [i, s] of spans.entries()) {
    if (s.noteRef !== undefined || s.invisible) continue
    const text = i === spans.length - 1 ? s.text.replace(/ +$/, '') : s.text
    if (text.length === 0) continue
    const size = s.fontSize * sizeScale
    const w = spanWidthPt(s, text, size, measure)
    if (w === null) return null
    const cps = [...text].length
    widthPt += w * (s.charScale ?? 1) + (s.charSpacingPt ?? 0) * cps
    codepoints += cps
  }
  return codepoints > 0 ? { widthPt, codepoints } : null
}

/**
 * Tighten one block in place so every line fits `availWidthPt` (the width
 * the rebuilt paragraph really gets) when re-rendered with output-font
 * advances. No-op when everything already fits — the common case: the
 * output family IS the source family and the column kept its width.
 */
export function preflightFitBlock(
  block: TextBlock,
  availWidthPt: number,
  measure: SpanMeasurer = systemFontMeasurer,
  opts: { strict?: boolean } = {},
): void {
  if (availWidthPt <= 0 || !measurable(block)) return
  const fontSize = Math.max(...block.lines.flatMap((l) => l.spans.map((s) => s.fontSize)), 1)
  const spacingCapPt = SPACING_CAP_EMS * fontSize
  const floorPt = Math.max(SHRINK_FLOOR_PT, SHRINK_FLOOR_RATIO * fontSize)
  // zone cells (P22 A) rebuild exact per-line geometry: the renderer wraps
  // strictly at the available width, so the anti-oversqueeze slack that suits
  // reflowable justified prose would let a measured display line wrap
  const overflowRatio = opts.strict ? 1 : OVERFLOW_RATIO
  const overflowAbsPt = opts.strict ? 0.25 : OVERFLOW_ABS_PT
  // compress to just under the wrap edge: the renderer's own advances can
  // run ~1% wider than the measured ones on substituted faces (strict mode:
  // zone cells whose exact line geometry depends on it)
  const targetRatio = opts.strict ? 0.99 : 0.998

  const availOf = (i: number): number =>
    (availWidthPt - (i === 0 ? Math.max(0, block.firstLineIndentPt) : 0)) * targetRatio

  /** spacing (pt/char, ≤0) that makes every line fit at `sizeScale`; null = unmeasurable */
  const requiredSpacing = (sizeScale: number): number | null => {
    let spacing = 0
    for (const [i, line] of block.lines.entries()) {
      const m = measureLine(line.spans, measure, sizeScale)
      if (!m) return null
      const avail = availOf(i)
      if (m.widthPt <= avail * overflowRatio + overflowAbsPt) continue
      spacing = Math.min(spacing, (avail - m.widthPt) / m.codepoints)
    }
    return spacing
  }

  let steps = 0
  let spacing = requiredSpacing(1)
  if (spacing === null || spacing === 0) return
  // one-twip grace on the cap: the tightened fit target must not tip a
  // borderline "spacing alone covers it" case into font shrinking
  while (
    spacing < -spacingCapPt - 1 / PT_TO_TWIPS &&
    steps < SHRINK_MAX_STEPS &&
    fontSize - (steps + 1) * SHRINK_STEP_PT >= floorPt
  ) {
    steps++
    const next = requiredSpacing((fontSize - steps * SHRINK_STEP_PT) / fontSize)
    if (next === null) return
    spacing = next
  }
  spacing = Math.max(spacing, -spacingCapPt)
  // snap to the w:spacing twip grid rounding DOWN: a -0.4-twip intent would
  // round to 0 at emission and the line would wrap after all
  spacing = Math.floor(spacing * PT_TO_TWIPS) / PT_TO_TWIPS

  const shrinkPt = steps * SHRINK_STEP_PT
  const scale = (fontSize - shrinkPt) / fontSize
  for (const line of block.lines) {
    for (const span of line.spans) {
      if (span.noteRef !== undefined) continue
      if (shrinkPt > 0) span.fontSize = Math.max(floorPt, span.fontSize * scale)
      if (spacing < -0.01) span.charSpacingPt = (span.charSpacingPt ?? 0) + spacing
    }
  }
}
