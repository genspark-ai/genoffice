import { describe, expect, it } from 'vitest'
import {
  printPaperOption,
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

  // Chromium prints on Letter unless the job states the document's own paper,
  // and a sheet taller than the paper is split in two — the printed page loses
  // its bottom and the page order goes with it (genoffice#2563)
  it('states the document paper to Electron in microns', () => {
    // A4: 11906 x 16838 twips = 210mm x 297mm
    expect(printPaperOption({ widthTwips: 11_906, heightTwips: 16_838 })).toEqual({
      pageSize: {
        width: Math.round((11_906 * 25_400) / 1440),
        height: Math.round((16_838 * 25_400) / 1440),
      },
    })
    const paper = printPaperOption({ widthTwips: 11_906, heightTwips: 16_838 }).pageSize!
    expect(paper.width).toBe(210_009)
    expect(paper.height).toBe(297_004)
    // Letter (8.5in x 11in) is what Chromium prints on when the job says nothing
    expect(paper.width).not.toBe(Math.round((12_240 * 25_400) / 1440))
    expect(paper.height).not.toBe(Math.round((15_840 * 25_400) / 1440))
  })

  it('states no paper for an absent or malformed one', () => {
    expect(printPaperOption(undefined)).toEqual({})
    expect(printPaperOption(null)).toEqual({})
    expect(printPaperOption({})).toEqual({})
    expect(printPaperOption({ widthTwips: 0, heightTwips: 16_838 })).toEqual({})
    expect(printPaperOption({ widthTwips: 11_906, heightTwips: 'x' })).toEqual({})
  })
})
