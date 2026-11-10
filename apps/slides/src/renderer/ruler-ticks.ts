/** Ruler tick math (PowerPoint-style): 0 sits at the slide center and values
 * grow toward both edges, in inches or centimeters. Pure module so the tick
 * layout is unit-testable without the DOM. */

export type RulerUnit = 'in' | 'cm'

/** Slide px per unit at zoom 1 (FIT_WIDTH viewport: 1280px ↔ 13.33in ↔ 96dpi) */
export const PX_PER_UNIT: Record<RulerUnit, number> = { in: 96, cm: 96 / 2.54 }

export type RulerTick = {
  /** ruler-local CSS px */
  pos: number
  /** unit distance from the slide center, absolute value; null = minor tick */
  label: string | null
}

/** Smallest label step (whole units) whose on-screen distance stays readable */
const LABEL_STEPS = [1, 2, 5, 10, 20, 50, 100]

export function computeRulerTicks({
  rulerLen,
  origin,
  slideLen,
  zoom,
  unit,
  minLabelPx = 28,
}: {
  /** visible ruler strip length, CSS px */
  rulerLen: number
  /** ruler-local CSS px where slide coordinate 0 falls (may be negative when scrolled) */
  origin: number
  /** slide length in unscaled slide px */
  slideLen: number
  zoom: number
  unit: RulerUnit
  /** minimum CSS px between two labeled ticks */
  minLabelPx?: number
}): RulerTick[] {
  const unitPx = PX_PER_UNIT[unit] * zoom
  if (!(unitPx > 0) || !(rulerLen > 0)) return []
  const step = LABEL_STEPS.find((s) => s * unitPx >= minLabelPx) ?? LABEL_STEPS.at(-1)!
  const half = step / 2
  const center = origin + (slideLen / 2) * zoom
  // Cover the whole visible strip, slide edges included and beyond (PPT's ruler
  // spans the work area, not just the page)
  const vFrom = Math.ceil((0 - center) / unitPx / half) * half
  const vTo = (rulerLen - center) / unitPx
  const ticks: RulerTick[] = []
  // Iterate integer multiples of `half` to dodge float drift on long strips
  for (let n = Math.round(vFrom / half); n * half <= vTo; n++) {
    const v = n * half
    const pos = center + v * unitPx
    if (pos < 0 || pos > rulerLen) continue
    const isLabel = n % 2 === 0
    ticks.push({ pos, label: isLabel ? String(Math.abs(v)) : null })
  }
  return ticks
}

/** Guide-bubble value: signed slide position (0..1 fraction) → distance from
 * the slide center in the ruler unit, e.g. "1.25 cm" (unsigned, like
 * PowerPoint's guide drag readout). */
export function formatRulerValue(frac: number, slideLen: number, unit: RulerUnit): string {
  const v = Math.abs((frac - 0.5) * slideLen) / PX_PER_UNIT[unit]
  const txt = v.toFixed(2).replace(/\.?0+$/, '') || '0'
  return `${txt} ${unit}`
}
