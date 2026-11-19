import { describe, expect, it } from 'vitest'
import {
  MARGIN_PRESETS,
  PAPER_SIZES,
  paperSizeCaption,
  paperSizeKeyOf,
} from '../src/renderer/components/ribbon-layout-tab'
import { formatLength } from '../src/renderer/units'

const cm = (twips: number) => formatLength(twips, 'cm')

describe('paperSizeCaption', () => {
  it('prints both sides in the preferred unit', () => {
    expect(paperSizeCaption(11906, 16838, cm)).toBe('21 cm \u00d7 29.7 cm')
    expect(paperSizeCaption(12240, 15840, cm)).toBe('21.59 cm \u00d7 27.94 cm')
    expect(paperSizeCaption(12240, 20160, cm)).toBe('21.59 cm \u00d7 35.56 cm')
    expect(paperSizeCaption(10319, 14572, cm)).toBe('18.2 cm \u00d7 25.7 cm')
    const inch = (twips: number) => formatLength(twips, 'in')
    expect(paperSizeCaption(12240, 15840, inch)).toBe('8.5" \u00d7 11"')
  })

  it('honors the translated unit', () => {
    const local = (twips: number) => formatLength(twips, 'cm', '\u5398\u7c73')
    expect(paperSizeCaption(11906, 16838, local)).toBe('21 \u5398\u7c73 \u00d7 29.7 \u5398\u7c73')
  })
})

describe('PAPER_SIZES', () => {
  it('lists Word\u2019s Size gallery in Word\u2019s order', () => {
    expect(PAPER_SIZES.map((p) => p.name)).toEqual([
      'Letter',
      'Legal',
      'Executive',
      'A3',
      'A4',
      'A5',
      'B4 (JIS)',
      'B5 (JIS)',
      'Tabloid',
      'Statement',
      'Envelope #10',
      'Envelope DL',
      'Envelope C5',
    ])
  })

  it('stores portrait twips matching the sheet dimensions', () => {
    const byName = Object.fromEntries(PAPER_SIZES.map((p) => [p.name, p]))
    expect(paperSizeCaption(byName.Executive.w, byName.Executive.h, cm)).toBe(
      '18.42 cm \u00d7 26.67 cm',
    )
    expect(paperSizeCaption(byName.A3.w, byName.A3.h, cm)).toBe('29.7 cm \u00d7 42 cm')
    expect(paperSizeCaption(byName['Envelope DL'].w, byName['Envelope DL'].h, cm)).toBe(
      '11 cm \u00d7 22 cm',
    )
    for (const p of PAPER_SIZES) expect(p.w).toBeLessThan(p.h)
    expect(new Set(PAPER_SIZES.map((p) => p.key)).size).toBe(PAPER_SIZES.length)
  })

  it('recognises the current page in either orientation and tells Letter from Legal', () => {
    expect(paperSizeKeyOf(12240, 15840)).toBe('letter')
    expect(paperSizeKeyOf(15840, 12240)).toBe('letter')
    expect(paperSizeKeyOf(12240, 20160)).toBe('legal')
    expect(paperSizeKeyOf(11907, 16839)).toBe('a4')
    expect(paperSizeKeyOf(10000, 10000)).toBeNull()
  })
})

describe('MARGIN_PRESETS', () => {
  it('adds Mirrored (inside 3.18 cm) and Office 2003 Default after Wide', () => {
    expect(MARGIN_PRESETS.map((m) => m.key)).toEqual([
      'normal',
      'narrow',
      'moderate',
      'wide',
      'mirrored',
      'office2003',
    ])
    const mirrored = MARGIN_PRESETS.find((m) => m.key === 'mirrored')!
    expect(mirrored.mirror).toBe(true)
    expect([mirrored.top, mirrored.bottom, mirrored.left, mirrored.right]).toEqual([
      1440, 1440, 1800, 1440,
    ])
    const office = MARGIN_PRESETS.find((m) => m.key === 'office2003')!
    expect(office.mirror).toBeUndefined()
    expect([office.top, office.bottom, office.left, office.right]).toEqual([1440, 1440, 1800, 1800])
  })
})
