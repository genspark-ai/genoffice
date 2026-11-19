// Ruler tick math: PowerPoint-style center-zero labels, adaptive density under
// zoom, cm/inch units, and coverage of the work area beyond the slide edges.
import { describe, expect, it } from 'vitest'
import { computeRulerTicks, formatRulerValue, PX_PER_UNIT } from '../src/renderer/ruler-ticks'

// 13.33in slide at zoom 1: 1280px wide, ruler exactly the slide, origin at 0
const SLIDE_W = 1280

describe('computeRulerTicks', () => {
  it('places 0 at the slide center and mirrors labels outward', () => {
    const ticks = computeRulerTicks({
      rulerLen: SLIDE_W,
      origin: 0,
      slideLen: SLIDE_W,
      zoom: 1,
      unit: 'in',
    })
    const zero = ticks.find((t) => t.label === '0')
    expect(zero).toBeDefined()
    expect(zero!.pos).toBeCloseTo(SLIDE_W / 2)
    // symmetric: a "2" on each side of the center
    const twos = ticks.filter((t) => t.label === '2')
    expect(twos).toHaveLength(2)
    expect(twos[0].pos + twos[1].pos).toBeCloseTo(SLIDE_W)
  })

  it('emits unlabeled minor ticks halfway between labels', () => {
    const ticks = computeRulerTicks({
      rulerLen: SLIDE_W,
      origin: 0,
      slideLen: SLIDE_W,
      zoom: 1,
      unit: 'in',
    })
    const center = SLIDE_W / 2
    const minor = ticks.find((t) => Math.abs(t.pos - (center + 0.5 * 96)) < 0.01)
    expect(minor).toBeDefined()
    expect(minor!.label).toBeNull()
  })

  it('widens the label step when zoomed out (no overlapping labels)', () => {
    const ticks = computeRulerTicks({
      rulerLen: SLIDE_W * 0.25,
      origin: 0,
      slideLen: SLIDE_W,
      zoom: 0.25,
      unit: 'in',
    })
    // 96px/in * 0.25 = 24px < 28px min → step 2in; only even labels appear
    const labels = ticks.filter((t) => t.label != null).map((t) => Number(t.label))
    expect(labels.length).toBeGreaterThan(0)
    for (const v of labels) expect(v % 2).toBe(0)
  })

  it('uses cm spacing for the cm unit', () => {
    const ticks = computeRulerTicks({
      rulerLen: SLIDE_W,
      origin: 0,
      slideLen: SLIDE_W,
      zoom: 1,
      unit: 'cm',
    })
    const zero = ticks.find((t) => t.label === '0')!
    const one = ticks
      .filter((t) => t.label === '1')
      .sort((a, b) => a.pos - b.pos)
      .at(-1)!
    expect(one.pos - zero.pos).toBeCloseTo(PX_PER_UNIT.cm)
  })

  it('keeps ticking beyond the slide edges across the work area', () => {
    // slide starts 200px into the ruler strip → ticks must exist before it
    const ticks = computeRulerTicks({
      rulerLen: SLIDE_W + 400,
      origin: 200,
      slideLen: SLIDE_W,
      zoom: 1,
      unit: 'in',
    })
    expect(ticks.some((t) => t.pos < 200)).toBe(true)
    expect(ticks.some((t) => t.pos > 200 + SLIDE_W)).toBe(true)
  })

  it('returns nothing for degenerate inputs', () => {
    expect(
      computeRulerTicks({ rulerLen: 0, origin: 0, slideLen: SLIDE_W, zoom: 1, unit: 'in' }),
    ).toEqual([])
    expect(
      computeRulerTicks({ rulerLen: 100, origin: 0, slideLen: SLIDE_W, zoom: 0, unit: 'in' }),
    ).toEqual([])
  })
})

describe('formatRulerValue', () => {
  it('reports unsigned distance from the slide center', () => {
    expect(formatRulerValue(0.5, SLIDE_W, 'in')).toBe('0 in')
    // full left edge of a 13.33in slide = 6.67in from center
    expect(formatRulerValue(0, SLIDE_W, 'in')).toBe('6.67 in')
    expect(formatRulerValue(1, SLIDE_W, 'in')).toBe('6.67 in')
  })

  it('trims trailing zeros', () => {
    // 0.25 of a 1280px slide → 320px from center → 320/96 = 3.33in
    expect(formatRulerValue(0.75, SLIDE_W, 'in')).toBe('3.33 in')
    const cm = formatRulerValue(1, SLIDE_W, 'cm')
    expect(cm.endsWith(' cm')).toBe(true)
  })
})
