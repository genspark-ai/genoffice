/**
 * Advance-width contract of the bundled KR subsets (fonts/README.md):
 * hangul 1.0em; serif digits/space follow Batang; sans Basic Latin follows
 * Malgun Gothic (tools/normalize-kr-sans-hmtx.py). A regenerated woff2 that
 * loses the normalization would silently shift Korean line breaks vs Word.
 */
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { advanceEm, readWoff2 } from './helpers/woff2-metrics'

const FONTS = join(__dirname, '../src/renderer/fonts')
const sans = readWoff2(join(FONTS, 'GenOfficeSansKR-Regular-subset.woff2'))
const serif = readWoff2(join(FONTS, 'GenOfficeSerifKR-Regular-subset.woff2'))

describe('GenOffice Sans KR (Malgun-normalized)', () => {
  it('hangul syllables and compatibility jamo stay 1.0em', () => {
    for (const cp of [0xac00, 0xae4e, 0xd558, 0x3131]) {
      expect(advanceEm(sans, cp), `U+${cp.toString(16)}`).toBe(1)
    }
  })

  it('space and digits match Malgun Gothic', () => {
    expect(advanceEm(sans, 0x20)).toBeCloseTo(0.352, 3)
    expect(advanceEm(sans, 0xa0)).toBeCloseTo(0.352, 3)
    for (let cp = 0x30; cp <= 0x39; cp++) {
      expect(advanceEm(sans, cp), `digit ${cp - 0x30}`).toBeCloseTo(0.551, 3)
    }
  })

  it('Basic Latin letters match Malgun Gothic', () => {
    const expected: Record<string, number> = {
      A: 0.658,
      J: 0.36,
      M: 0.917,
      W: 0.954,
      a: 0.52,
      i: 0.246,
      m: 0.88,
      '.': 0.219,
      ',': 0.219,
      '(': 0.305,
      ')': 0.305,
      '/': 0.396,
      '%': 0.836,
    }
    for (const [ch, adv] of Object.entries(expected)) {
      expect(advanceEm(sans, ch.codePointAt(0)!), ch).toBeCloseTo(adv, 3)
    }
  })

  it('fullwidth forms keep the 1.0em CJK advance', () => {
    expect(advanceEm(sans, 0x3000)).toBe(1)
    expect(advanceEm(sans, 0xff10)).toBe(1)
  })
})

describe('GenOffice Serif KR (Batang-normalized)', () => {
  it('hangul 1.0em, digits 0.596em, space 0.333em', () => {
    expect(advanceEm(serif, 0xac00)).toBe(1)
    expect(advanceEm(serif, 0x3131)).toBe(1)
    expect(advanceEm(serif, 0x20)).toBeCloseTo(0.333, 3)
    for (let cp = 0x30; cp <= 0x39; cp++) {
      expect(advanceEm(serif, cp), `digit ${cp - 0x30}`).toBeCloseTo(0.596, 3)
    }
  })
})

describe('bundled SC subset fullwidth coverage (TC serif chain shim)', () => {
  it('U+FF0D/FF0F/FF3C/FF3F/FF5E map to 1.0em glyphs', () => {
    const sc = readWoff2(join(FONTS, 'NotoSerifCJKsc-Regular-subset.woff2'))
    for (const cp of [0xff0d, 0xff0f, 0xff3c, 0xff3f, 0xff5e]) {
      expect(advanceEm(sc, cp), `U+${cp.toString(16)}`).toBe(1)
    }
  })
})
