import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { fileAdvanceWidths, isWoff2, woff2ToSfnt } from '../src'

const DOCS_FONTS = resolve(__dirname, '../../../apps/docs/src/renderer/fonts')
const KR = resolve(DOCS_FONTS, 'GenOfficeSansKR-Regular-subset.woff2')
const THAI = resolve(DOCS_FONTS, 'GenOfficeSarabun-Regular.woff2')
const CARLITO = resolve(__dirname, '../../ui/src/fonts/Carlito-Regular.ttf')

describe('woff2ToSfnt', () => {
  it('keeps the untransformed metric tables of a bundled subset', () => {
    const data = readFileSync(KR)
    expect(isWoff2(data)).toBe(true)
    const sfnt = woff2ToSfnt(data)!
    const tags = new Set<string>()
    for (let t = 0; t < sfnt.readUInt16BE(4); t++)
      tags.add(sfnt.toString('latin1', 12 + 16 * t, 16 + 16 * t))
    for (const tag of ['cmap', 'head', 'hhea', 'hmtx']) expect(tags.has(tag)).toBe(true)
  })

  it('is not a woff2: plain sfnt and garbage return null', () => {
    expect(woff2ToSfnt(readFileSync(CARLITO))).toBeNull()
    expect(woff2ToSfnt(Buffer.from('wOF2 but far too short'))).toBeNull()
  })

  it('fileAdvanceWidths reads a .woff2 face', () => {
    // Noto Sans KR subset: proportional Latin, full-width Hangul, no Han
    const [a, han, hangul] = fileAdvanceWidths(KR, 'a中한', 1000)!
    expect(a).toBeGreaterThan(8000)
    expect(a).toBeLessThan(13000)
    expect(hangul).toBe(20000)
    expect(han).toBeNaN()
    const [thai] = fileAdvanceWidths(THAI, 'ก', 1000)!
    expect(thai).toBeGreaterThan(5000)
    expect(thai).toBeLessThan(12000)
  })
})
