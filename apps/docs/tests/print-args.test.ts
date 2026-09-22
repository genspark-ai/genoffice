import { describe, expect, it } from 'vitest'
import { printScaleOption, validPrintDim, validPrintScale } from '../src/main/print-args'

describe('print arg validators', () => {
  it('accepts real page sizes and default scale', () => {
    expect(validPrintDim(12240)).toBe(true) // Letter 8.5in
    expect(validPrintDim(15840)).toBe(true) // A4 height-ish
    expect(validPrintScale(undefined)).toBe(true)
    expect(validPrintScale(1)).toBe(true)
  })

  it('rejects non-finite and out-of-range geometry', () => {
    for (const v of [NaN, Infinity, -Infinity, 0, -100, 143, 72001, 1e12, '12240', null]) {
      expect(validPrintDim(v)).toBe(false)
    }
    for (const s of [NaN, Infinity, -Infinity, 0, 0.05, 5.1, 1e9, '2']) {
      expect(validPrintScale(s)).toBe(false)
    }
  })

  it('drops non-finite scales from the Chromium option objects', () => {
    expect(printScaleOption(1)).toEqual({})
    expect(printScaleOption(2)).toEqual({ scale: 2 })
    expect(printScaleOption(Infinity)).toEqual({})
    expect(printScaleOption(NaN)).toEqual({})
  })
})
