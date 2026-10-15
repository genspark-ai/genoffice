import { describe, expect, it } from 'vitest'
import {
  printScaleOption,
  validPrintDim,
  validPrintGeometry,
  validPrintScale,
} from '../src/main/print-args'

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

  // regression: docs:export-pdf used to skip this guard entirely, so a doc with
  // <w:pgSz w:w="2000000000"/> reached printToPDF as a ~1,388,889-inch page
  it('validPrintGeometry gates the export-pdf geometry+scale combination', () => {
    // legit Letter portrait passes
    expect(validPrintGeometry(12240, 15840)).toBe(true)
    expect(validPrintGeometry(12240, 15840, 1.5)).toBe(true)
    expect(validPrintGeometry(12240, 15840, undefined)).toBe(true)
    // hostile pgSz from a doc (w:w="2000000000") is rejected before printToPDF
    expect(validPrintGeometry(2_000_000_000, 15840)).toBe(false)
    expect(validPrintGeometry(12240, 2_000_000_000)).toBe(false)
    // NaN/Infinity/strings and bad scale rejected too
    expect(validPrintGeometry(NaN, 15840)).toBe(false)
    expect(validPrintGeometry(12240, Infinity)).toBe(false)
    expect(validPrintGeometry('12240', 15840)).toBe(false)
    expect(validPrintGeometry(12240, 15840, 99)).toBe(false)
  })
})
