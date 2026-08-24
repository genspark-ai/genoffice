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

  it('never inflates an uncalibrated fallback measurement', () => {
    expect(excelWidthScale('Calibri', 11, () => 6)).toBe(1)
  })

  it('inflates a calibrated narrower substitute up to the GDI width', () => {
    // 96%-Carlito digit at 11pt ≈ 7.14px vs Excel's GDI 8px — Excel hashes
    // cells our substitute would still fit (prod_016).
    expect(excelWidthScale('Aptos Narrow', 11, () => 7.14, true)).toBeCloseTo(8 / 7.14)
    // With the genuine font installed (no substitute registered), never
    // inflate past the live canvas measurement.
    expect(excelWidthScale('Aptos Narrow', 11, () => 7.14, false)).toBe(1)
  })

  it('leaves unknown or unset families alone', () => {
    expect(excelWidthScale('Arial', 11, () => 8.25)).toBe(1)
    expect(excelWidthScale(undefined, 11, () => 8.25)).toBe(1)
  })

  it('scales alias-substituted Korean defaults even when the alias resolves', () => {
    // The styles.css alias maps the family onto Apple SD Gothic Neo, whose
    // digits measure ~10px at 11pt vs the GDI 7px.
    expect(excelWidthScale('맑은 고딕', 11, () => 9.97)).toBeCloseTo(7 / 9.97)
    expect(excelWidthScale('Malgun Gothic', 11, () => 9.97)).toBeCloseTo(7 / 9.97)
  })

  it('keeps a borderline cell from spuriously hashing', () => {
    // 10 chars * 8px = 80 > 75 raw, but 80 * 0.9 = 72 fits.
    expect(overflowHashes('2/22/2016 ', 80, measure, 0.9)).toBeNull()
    expect(overflowHashes('2/22/2016 ', 80, measure)).toBe('#########')
  })

  it('tolerates overflow within the GDI measurement noise band', () => {
    // 10 chars * 8px = 80 > 77 available, but 77 * 1.05 = 80.85 fits —
    // Excel's own (GDI) metrics would show the value, so clip, not hash.
    expect(overflowHashes('2026/8/30x', 82, measure)).toBeNull()
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
