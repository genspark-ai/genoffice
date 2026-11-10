import { describe, expect, it } from 'vitest'

import { scatterAxisBounds, valueAxisScale } from '../src/domain/chart-visual'

describe('axis tick generation', () => {
  it('ticks a fine majorUnit all the way to the axis maximum', () => {
    // 0..10 with c:majorUnit 0.1 implies 101 ticks; a fixed 25-iteration
    // walk stopped at 2.4, leaving the axis short of its own data.
    const scale = valueAxisScale(10, { min: 0, max: 10, majorUnit: 0.1 })
    expect(scale.ticks).toHaveLength(101)
    expect(scale.ticks[0]).toBe(0)
    expect(scale.ticks[50]).toBe(5)
    expect(scale.ticks.at(-1)).toBe(10)
  })

  it('keeps the calibrated auto-scale tick counts', () => {
    // Auto paths stay at <= 10 intervals, so the old cap never bound.
    expect(valueAxisScale(18).ticks).toEqual([0, 2, 4, 6, 8, 10, 12, 14, 16, 18, 20])
    expect(valueAxisScale(877).ticks).toEqual([
      0, 100, 200, 300, 400, 500, 600, 700, 800, 900, 1000,
    ])
  })

  it('ticks a fine scatter majorUnit up to the axis maximum', () => {
    const axis = scatterAxisBounds([0, 1], { min: 0, max: 1, majorUnit: 0.25 })
    expect(axis.ticks).toEqual([0, 0.25, 0.5, 0.75, 1])
  })

  it('falls back to the bounds for a zero-width unit', () => {
    expect(valueAxisScale(10, { min: 0, max: 10, majorUnit: 0 }).ticks).toEqual([0, 10])
  })
})
