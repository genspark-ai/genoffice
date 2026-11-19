import { describe, expect, it } from 'vitest'
import { logMinors, logTicks } from '../src/build-chart'

describe('logTicks finite-range guards', () => {
  it('terminates on a data max near the number ceiling instead of spinning forever', () => {
    // base ** e overflows to Infinity for a data point near 1e308; the
    // unguarded loop then compared Infinity <= Infinity forever (one hostile
    // number froze the slide for good). Measured: >100k iterations and
    // climbing before the guard; now it returns finite ticks.
    const t0 = performance.now()
    const { min, max, ticks } = logTicks(1.5e308, 1.5e308, undefined, undefined, 10)
    expect(performance.now() - t0).toBeLessThan(250)
    expect(Number.isFinite(min)).toBe(true)
    expect(Number.isFinite(max)).toBe(true)
    expect(ticks.length).toBeLessThanOrEqual(200)
    for (const t of ticks) expect(Number.isFinite(t)).toBe(true)
  })

  it('keeps the sane-case tick set unchanged', () => {
    const { ticks } = logTicks(1, 1e7, undefined, undefined, 10)
    expect(ticks).toEqual([1, 10, 100, 1000, 10000, 100000, 1000000, 10000000])
  })

  it('clamps a non-finite explicit max and a non-finite positive minimum', () => {
    const a = logTicks(1, 1e6, undefined, Infinity, 10)
    expect(Number.isFinite(a.max)).toBe(true)
    expect(Number.isFinite(a.min)).toBe(true)
    // posDataMin of Infinity would make lo itself infinite; the guard falls
    // back to the base so the axis still renders.
    const b = logTicks(Number.POSITIVE_INFINITY, 1e6, undefined, undefined, 10)
    expect(Number.isFinite(b.min)).toBe(true)
    expect(b.ticks.length).toBeGreaterThan(0)
  })
})

describe('logMinors finite-range guards', () => {
  it('returns no minors for an infinite range instead of looping forever', () => {
    const t0 = performance.now()
    expect(logMinors(1e308, Number.POSITIVE_INFINITY, 10)).toEqual([])
    expect(performance.now() - t0).toBeLessThan(100)
  })

  it('keeps producing minors for a sane range', () => {
    const minors = logMinors(1, 1000, 10)
    expect(minors.length).toBeGreaterThan(0)
    expect(minors).toContain(20)
    expect(minors).toContain(900)
  })
})
