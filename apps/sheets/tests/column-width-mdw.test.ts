import { WrapStrategy } from '@univerjs/core'
import { afterEach, describe, expect, it } from 'vitest'

import {
  getWorkbookMdw,
  pixelsToCharacterWidth,
  setWorkbookMdw,
} from '../src/renderer/app-constants'
import { generalCharBudget } from '../src/renderer/numfmt-fix'
import { characterWidthToPixels, toUniverStyle } from '../src/renderer/univer-sync'

afterEach(() => setWorkbookMdw(7))

describe('workbook MDW', () => {
  it('defaults to Calibri 11 (7px) and keeps the historical conversion', () => {
    expect(getWorkbookMdw()).toBe(7)
    expect(characterWidthToPixels(10.6640625)).toBe(80)
  })

  it('widens columns for a Verdana-10 workbook (MDW 8)', () => {
    // DateFormatTests.xlsx col C: width 34.832 chars. Excel renders 283.7px
    // (width x 8 + 5); the hardcoded 7 yielded 249px and wrapped a line early.
    setWorkbookMdw(8)
    expect(characterWidthToPixels(34.83203125)).toBe(284)
  })

  it('keeps the General digit budget on the same MDW', () => {
    // A Verdana-10 column imported as 40 chars must still budget 40 digits,
    // not (40*8)/7.
    setWorkbookMdw(8)
    expect(generalCharBudget(characterWidthToPixels(40))).toBe(40)
  })

  it('keeps both conversion directions on the same MDW', () => {
    setWorkbookMdw(8)
    const px = characterWidthToPixels(12)
    expect(Math.abs(pixelsToCharacterWidth(px) - 12)).toBeLessThan(0.05)
  })

  it('clamps nonsense MDW values back to 7', () => {
    setWorkbookMdw(0)
    expect(getWorkbookMdw()).toBe(7)
    setWorkbookMdw(Number.NaN)
    expect(getWorkbookMdw()).toBe(7)
  })
})

describe('toUniverStyle wrap resolution', () => {
  const base = {
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
  }

  it('emits WRAP for wrapping styles and an explicit OVERFLOW otherwise', () => {
    expect(toUniverStyle({ ...base, wrapText: true } as never).tb).toBe(WrapStrategy.WRAP)
    // A resolved non-wrap cell xf must override a WRAP column style at
    // compose time (sample 60384: col style wraps, A1 explicitly does not).
    expect(toUniverStyle({ ...base, wrapText: false } as never).tb).toBe(WrapStrategy.OVERFLOW)
  })
})
