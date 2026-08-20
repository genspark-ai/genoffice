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
 * Scope is deliberately narrow: LTR blocks of Latin-ish text whose families
 * the measurer resolves. CJK reflows gracefully (any character may wrap) and
 * complex scripts need shaping the measurer does not do.
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

const MEASURABLE_SCRIPTS = new Set(['latin', 'common'])

const measurable = (block: TextBlock): boolean =>
  block.dir === 'ltr' &&
  block.lines.length >= 1 &&
  block.tocEntry === undefined &&
  block.list === undefined &&
  block.cardId === undefined &&
  block.lines.every((line) =>
    line.spans.every(
      (s) =>
        s.noteRef !== undefined ||
        s.invisible === true ||
        (MEASURABLE_SCRIPTS.has(s.script) && !s.text.includes('\t') && s.fontFamily !== ''),
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
    const w = measure(text, s.fontFamily, size, s.bold, s.italic)
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
