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

const PRIMARY_NAME_IDS = new Set([1, 2, 3, 4, 6, 16, 17, 18, 20, 21, 22, 25])

function decodeUtf16Be(value: Buffer): string {
  const swapped = Buffer.alloc(value.length)
  for (let i = 0; i + 1 < value.length; i += 2) {
    swapped[i] = value[i + 1]
    swapped[i + 1] = value[i]
  }
  return swapped.toString('utf16le')
}

function primaryFontNames(font: ReturnType<typeof readWoff2>): Array<{
  nameId: number
  value: string
}> {
  const table = font.tables.get('name')
  if (!table) throw new Error('font has no name table')
  const count = table.readUInt16BE(2)
  const storageOffset = table.readUInt16BE(4)
  const names = []
  for (let i = 0; i < count; i++) {
    const recordOffset = 6 + i * 12
    const platformId = table.readUInt16BE(recordOffset)
    const nameId = table.readUInt16BE(recordOffset + 6)
    if (!PRIMARY_NAME_IDS.has(nameId)) continue
    const length = table.readUInt16BE(recordOffset + 8)
    const offset = table.readUInt16BE(recordOffset + 10)
    const raw = table.subarray(storageOffset + offset, storageOffset + offset + length)
    const value = platformId === 0 || platformId === 3 ? decodeUtf16Be(raw) : raw.toString('latin1')
    names.push({ nameId, value })
  }
  return names
}

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

describe('GenOffice Gothic KR (real source metrics, unmodified)', () => {
  const gothicKr = readWoff2(join(FONTS, 'GenOfficeGothicKR-Regular-subset.woff2'))

  it('keeps the M3 probe truth: hangul 0.94em, space 0.28em, digits 0.606em', () => {
    for (const cp of [0xac00, 0xd55c, 0xae00]) {
      expect(advanceEm(gothicKr, cp), `U+${cp.toString(16)}`).toBeCloseTo(0.94, 3)
    }
    expect(advanceEm(gothicKr, 0x20)).toBeCloseTo(0.28, 3)
    for (let cp = 0x30; cp <= 0x39; cp++) {
      expect(advanceEm(gothicKr, cp), `digit ${cp - 0x30}`).toBeCloseTo(0.606, 3)
    }
  })

  it('covers KS X 1001 syllables and Basic Latin', () => {
    expect(advanceEm(gothicKr, 0xd558)).toBeCloseTo(0.94, 3)
    expect(advanceEm(gothicKr, 0x41)).toBeGreaterThan(0)
  })

  it('keeps every primary font identifier clear of Reserved Font Names', () => {
    const names = primaryFontNames(gothicKr)
    expect(names.length).toBeGreaterThan(0)
    for (const { nameId, value } of names) {
      expect(value, `name ID ${nameId}`).not.toMatch(/nanum/i)
    }
    expect(names).toContainEqual({ nameId: 1, value: 'GenOffice Gothic KR' })
    expect(names).toContainEqual({ nameId: 6, value: 'GenOfficeGothicKR-Regular' })
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
