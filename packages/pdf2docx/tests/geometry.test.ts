import { describe, expect, it } from 'vitest'
import {
  approxEq,
  coverageRatio,
  intersectArea,
  median,
  overlapRatio,
  rectUnion,
  verticalOverlapRatio,
} from '../src/geometry'

describe('tolerance helpers', () => {
  it('approxEq honors the tolerance in both directions', () => {
    expect(approxEq(1.0, 1.05, 0.1)).toBe(true)
    expect(approxEq(1.0, 1.11, 0.1)).toBe(false)
    expect(approxEq(-3, -3.09, 0.1)).toBe(true)
  })

  it('median handles odd/even/empty lists', () => {
    expect(median([3, 1, 2])).toBe(2)
    expect(median([4, 1, 2, 3])).toBe(2.5)
    expect(median([])).toBe(0)
  })
})

describe('rect math', () => {
  const a = { x0: 0, y0: 0, x1: 10, y1: 10 }
  const b = { x0: 5, y0: 5, x1: 15, y1: 15 }

  it('union / intersection / overlap', () => {
    expect(rectUnion(a, b)).toEqual({ x0: 0, y0: 0, x1: 15, y1: 15 })
    expect(intersectArea(a, b)).toBe(25)
    expect(overlapRatio(a, b)).toBeCloseTo(0.25)
    expect(intersectArea(a, { x0: 20, y0: 20, x1: 30, y1: 30 })).toBe(0)
  })

  it('vertical overlap ratio uses the shorter rect', () => {
    const line = { x0: 0, y0: 0, x1: 100, y1: 10 }
    const sup = { x0: 50, y0: 6, x1: 55, y1: 14 } // 4 of its 8 units overlap
    expect(verticalOverlapRatio(sup, line)).toBeCloseTo(0.5)
  })
})

describe('coverageRatio (P29 E)', () => {
  it('full-page box covers everything', () => {
    expect(coverageRatio([{ x0: 0, y0: 0, x1: 600, y1: 800 }], 600, 800)).toBe(1)
  })
  it('half-page box covers about half', () => {
    const r = coverageRatio([{ x0: 0, y0: 0, x1: 600, y1: 400 }], 600, 800)
    expect(r).toBeGreaterThan(0.45)
    expect(r).toBeLessThan(0.56)
  })
  it('no boxes cover nothing', () => {
    expect(coverageRatio([], 600, 800)).toBe(0)
  })
})
