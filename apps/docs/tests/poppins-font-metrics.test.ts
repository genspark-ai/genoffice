/**
 * Contract of the bundled Poppins subsets (fonts/README.md): Word downloads
 * the real M365 cloud face and lays out with a 1.5em line box and Poppins'
 * own advances (probe 2026-09-01). A regenerated woff2 that normalizes either
 * would silently shift Poppins documents' line breaks and page count vs Word.
 */
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { advanceEm, readWoff2 } from './helpers/woff2-metrics'

const FONTS = join(__dirname, '../src/renderer/fonts')
const regular = readWoff2(join(FONTS, 'GenOfficePoppins-Regular-subset.woff2'))
const bold = readWoff2(join(FONTS, 'GenOfficePoppins-Bold-subset.woff2'))

describe('GenOffice Poppins (real-metric M365 cloud face)', () => {
  it('keeps the upstream 1.5em hhea line box', () => {
    for (const font of [regular, bold]) {
      const hhea = font.tables.get('hhea')!
      const ascent = hhea.readInt16BE(4)
      const descent = hhea.readInt16BE(6)
      const lineGap = hhea.readInt16BE(8)
      expect((ascent - descent + lineGap) / font.unitsPerEm).toBe(1.5)
    }
  })

  it('keeps upstream advances (space/digits, per weight)', () => {
    expect(advanceEm(regular, 0x20)).toBe(0.267)
    expect(advanceEm(regular, 0x30)).toBe(0.628)
    expect(advanceEm(bold, 0x20)).toBe(0.212)
    expect(advanceEm(bold, 0x30)).toBe(0.652)
  })

  it('covers Latin letters and typographic punctuation', () => {
    for (const cp of [0x41, 0x7a, 0xe9, 0x201c, 0x2014, 0x20ac]) {
      expect(regular.cmap.get(cp), `U+${cp.toString(16)}`).toBeDefined()
    }
  })
})
