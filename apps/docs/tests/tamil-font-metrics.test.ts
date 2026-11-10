/**
 * Advance-width contract of the bundled Tamil face (fonts/README.md): advances
 * rewritten to Word's Latha (tools/build-tamil-font.py). A regenerated woff2
 * that loses the normalization would drift Tamil line breaks ~27% vs Word
 * (M3 probe 2026-08-14: Tamil Sangam MN sentence R 0.728, space 0.39x).
 */
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { advanceEm, readWoff2 } from './helpers/woff2-metrics'

const FONTS = join(__dirname, '../src/renderer/fonts')
const tamil = readWoff2(join(FONTS, 'GenOfficeTamil-Regular.woff2'))

describe('GenOffice Tamil (Latha-normalized)', () => {
  it('space and digits match Latha', () => {
    expect(advanceEm(tamil, 0x20)).toBeCloseTo(0.578, 3)
    for (let cp = 0x30; cp <= 0x39; cp++) {
      expect(advanceEm(tamil, cp), `digit ${cp - 0x30}`).toBeCloseTo(0.5, 3)
    }
  })

  it('Tamil letters and matras match Latha', () => {
    const expected: Record<number, number> = {
      0x0b85: 1.133, // அ
      0x0b86: 1.352, // ஆ
      0x0b95: 0.797, // க
      0x0ba4: 0.836, // த
      0x0bae: 0.965, // ம
      0x0bbe: 0.617, // ா
      0x0bbf: 0.188, // ி
      0x0bc1: 0.496, // ு
    }
    for (const [cp, adv] of Object.entries(expected)) {
      expect(advanceEm(tamil, Number(cp)), `U+${Number(cp).toString(16)}`).toBeCloseTo(adv, 3)
    }
  })
})
