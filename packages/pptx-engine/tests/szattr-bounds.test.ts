import { describe, expect, it } from 'vitest'
import { clampInt, szAttr } from '../src/generate'

/**
 * Regression: szAttr used Math.max/min directly, which propagate NaN, so a
 * non-finite font size serialized as sz="NaN". It now lands on the 1pt lower
 * bound, matching clampInt's documented behavior.
 */
describe('font size attribute bounds', () => {
  it('clamps normal sizes to hundredths', () => {
    expect(szAttr(18)).toBe('1800')
    expect(szAttr(0.5)).toBe('100')
    expect(szAttr(5000)).toBe('400000')
  })

  it('lands non-finite sizes on the lower bound', () => {
    expect(szAttr(NaN)).toBe('100')
    expect(szAttr(Infinity)).toBe('100')
    expect(clampInt(NaN, 0, 100)).toBe(0)
  })
})
