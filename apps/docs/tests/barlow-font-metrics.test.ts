/**
 * Contract of the bundled Barlow subsets (fonts/README.md): Word downloads the
 * cloud face and lays a Barlow document out with its 1.2em line box and
 * Barlow's own advances (corpus 2026-11-19). A regenerated woff2 that
 * normalizes either would silently shift line breaks and page counts.
 */
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { advanceEm, readWoff2 } from './helpers/woff2-metrics'

const FONTS = join(__dirname, '../src/renderer/fonts')
const regular = readWoff2(join(FONTS, 'GenOfficeBarlow-Regular-subset.woff2'))
const bold = readWoff2(join(FONTS, 'GenOfficeBarlow-Bold-subset.woff2'))

describe('GenOffice Barlow (real-metric face)', () => {
  it('keeps the upstream 1.2em hhea line box', () => {
    for (const font of [regular, bold]) {
      const hhea = font.tables.get('hhea')!
      const ascent = hhea.readInt16BE(4)
      const descent = hhea.readInt16BE(6)
      const lineGap = hhea.readInt16BE(8)
      expect((ascent - descent + lineGap) / font.unitsPerEm).toBe(1.2)
    }
  })

  it('keeps upstream advances (space/digits/a, per weight)', () => {
    expect(advanceEm(regular, 0x20)).toBe(0.2)
    expect(advanceEm(regular, 0x30)).toBe(0.565)
    expect(advanceEm(regular, 0x61)).toBe(0.511)
    expect(advanceEm(bold, 0x20)).toBe(0.2)
    expect(advanceEm(bold, 0x30)).toBe(0.57)
  })

  it('covers Latin letters and typographic punctuation', () => {
    for (const cp of [0x41, 0x7a, 0xe9, 0x201c, 0x2014, 0x20ac]) {
      expect(regular.cmap.get(cp), `U+${cp.toString(16)}`).toBeDefined()
    }
  })
})
