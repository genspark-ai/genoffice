import { describe, expect, it } from 'vitest'

import {
  excelWidthScale,
  hasGdiDigitCalibration,
  hashFill,
  overflowHashes,
} from '../src/renderer/numfmt-fix'

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

  it('scales alias-substituted ja faces to their genuine digit widths', () => {
    // Hiragino substitute digits 0.657em (9.64px at 11pt); the genuine
    // MS PGothic is 0.5em (7.33px) and Yu Gothic 0.556em (8.16px).
    expect(excelWidthScale('ＭＳ Ｐゴシック', 11, () => 9.64)).toBeCloseTo(7.33 / 9.64, 2)
    expect(excelWidthScale('Yu Gothic', 11, () => 9.64)).toBeCloseTo(8.16 / 9.64, 2)
  })

  it('tolerates overflow within the GDI measurement noise band', () => {
    // 10 chars * 8px = 80 > 77 available, but 77 * 1.05 = 80.85 fits —
    // Excel's own (GDI) metrics would show the value, so clip, not hash.
    expect(overflowHashes('2026/8/30x', 82, measure)).toBeNull()
  })

  it('hashes a calibrated measurement past Excel’s digit-width slack (prod_023)', () => {
    // Genuine Yu Gothic measures at its GDI width (live probe: digit 8.14px,
    // "1,225" 36.34px) so the scale clamps to 1, but the family is
    // table-calibrated: 36.34 > (47−5) − 8.14 → Excel hashes.
    const m = (text: string): number => (text === '1,225' ? 36.34 : 8.14)
    expect(overflowHashes('1,225', 47, m, 1, true)).not.toBeNull()
    // The same numbers under the uncalibrated noise band keep the value.
    expect(overflowHashes('1,225', 47, m, 1, false)).toBeNull()
  })

  it('keeps a calibrated measurement inside the slack (prod_037)', () => {
    // MS PGothic date via the substitute: 91px, digit 9.64px, scale 0.76 →
    // true 69.2px vs limit (103−5) − 7.33 = 90.7 → the date shows.
    const m = (text: string): number => (text === '2017/08/22' ? 91 : 9.64)
    expect(overflowHashes('2017/08/22', 103, m, 0.76)).toBeNull()
  })

  it('limits digit-slack calibration to the substituted ja faces', () => {
    expect(hasGdiDigitCalibration('Yu Gothic')).toBe(true)
    expect(hasGdiDigitCalibration('ＭＳ Ｐゴシック')).toBe(true)
    expect(hasGdiDigitCalibration('Calibri')).toBe(false)
    expect(hasGdiDigitCalibration(undefined)).toBe(false)
  })

  it('scales substituted Cordia New back to its narrow print digits (prod_066)', () => {
    // Ref print: digits advance 4pt per 11pt = 5.33px; the Helvetica
    // fallback measures 8.15px, hashing amounts Excel fits.
    const scale = excelWidthScale('Cordia New', 11, () => 8.15)
    expect(scale).toBeCloseTo(5.333 / 8.15, 2)
    expect(hasGdiDigitCalibration('Cordia New')).toBe(true)
    const m = (text: string): number => (text === '2,867,943.93' ? 85.6 : 8.15)
    expect(overflowHashes('2,867,943.93', 84, m, scale, true)).toBeNull()
    // The unscaled fallback measurement tripped the hash rule.
    expect(overflowHashes('2,867,943.93', 84, m)).not.toBeNull()
  })

  it('treats the width-corrected Thai and JP aliases as calibrated substitutes', () => {
    // Every family with a size-adjusted alias carries its em-exact digit
    // width, so the corrected canvas digit lands on the table value (scale
    // ~1) and the digit-slack rule applies instead of the noise band.
    for (const family of [
      'MS UI Gothic',
      '游ゴシック体',
      'Yu Gothic UI',
      'メイリオ',
      'Meiryo',
      'Meiryo UI',
      'Angsana New',
      'AngsanaUPC',
      'TH SarabunPSK',
      'TH Sarabun New',
    ]) {
      expect(hasGdiDigitCalibration(family), family).toBe(true)
    }
    // Cordia New via the 65% Helvetica Neue sub-face: 0.556em × 0.65 ×
    // 14.67px = 5.30px against the 5.33px print digit.
    expect(excelWidthScale('Cordia New', 11, () => 5.3, true)).toBeCloseTo(5.333 / 5.3, 2)
    // MS UI Gothic shares MS PGothic's 0.5em digits (7.33px at 11pt).
    expect(excelWidthScale('MS UI Gothic', 11, () => 7.33, true)).toBeCloseTo(1, 2)
    // TH SarabunPSK 0.362em → 5.31px at 11pt.
    expect(excelWidthScale('TH SarabunPSK', 11, () => 5.31, true)).toBeCloseTo(1, 2)
  })

  it('lifts the >1 clamp only for a registered narrow substitute (prod_066)', () => {
    // A slightly-too-narrow substitute digit must still be inflated to the
    // GDI width so Excel's #### trips; the genuine font never inflates.
    expect(excelWidthScale('Cordia New', 11, () => 5.2, true)).toBeGreaterThan(1)
    expect(excelWidthScale('Cordia New', 11, () => 5.2, false)).toBe(1)
    expect(excelWidthScale('MS PGothic', 11, () => 7.1, true)).toBeGreaterThan(1)
    expect(excelWidthScale('MS PGothic', 11, () => 7.1, false)).toBe(1)
  })

  it('hashes a too-wide boolean label (prod_074)', () => {
    // Arial FALSE 42.2px in a 33px column: (33−5)×1.05 = 29.4 < 42.2.
    const m = (text: string): number => (text === 'FALSE' ? 42.2 : 8)
    expect(overflowHashes('FALSE', 33, m)).toBe('###')
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
