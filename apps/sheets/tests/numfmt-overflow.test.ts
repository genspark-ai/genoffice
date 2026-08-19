import { describe, expect, it } from 'vitest'

import { excelWidthScale, hashFill, overflowHashes } from '../src/renderer/numfmt-fix'

const measure = (text: string): number => text.length * 8

describe('overflowHashes', () => {
  it('returns null when the text fits the column', () => {
    expect(overflowHashes('$472.00', 85, measure)).toBeNull()
  })

  it('fills the available width with hashes on overflow', () => {
    // 16 chars * 8px = 128 > 85 - 5; fill = floor(80 / 8) = 10.
    expect(overflowHashes('2015-12-15 22:50', 85, measure)).toBe('##########')
  })

  it('keeps at least one hash in a sliver column', () => {
    expect(overflowHashes('99', 6, measure)).toBe('#')
  })

  it('ignores empty text', () => {
    expect(overflowHashes('', 85, measure)).toBeNull()
  })
})

describe('excelWidthScale', () => {
  it('scales substituted Calibri back to the GDI digit width', () => {
    // Helvetica digit at 11pt ≈ 8.25px vs Excel's 7px.
    expect(excelWidthScale('Calibri', 11, () => 8.25)).toBeCloseTo(7 / 8.25)
  })

  it('never inflates a measurement', () => {
    expect(excelWidthScale('Calibri', 11, () => 6)).toBe(1)
  })

  it('leaves unknown or unset families alone', () => {
    expect(excelWidthScale('Arial', 11, () => 8.25)).toBe(1)
    expect(excelWidthScale(undefined, 11, () => 8.25)).toBe(1)
  })

  it('keeps a borderline cell from spuriously hashing', () => {
    // 10 chars * 8px = 80 > 75 raw, but 80 * 0.9 = 72 fits.
    expect(overflowHashes('2/22/2016 ', 80, measure, 0.9)).toBeNull()
    expect(overflowHashes('2/22/2016 ', 80, measure)).toBe('#########')
  })
})

describe('hashFill', () => {
  it('fills regardless of the text', () => {
    expect(hashFill(85, measure)).toBe('##########')
  })

  it('returns null when the hash glyph cannot be measured', () => {
    expect(hashFill(85, () => 0)).toBeNull()
  })
})
