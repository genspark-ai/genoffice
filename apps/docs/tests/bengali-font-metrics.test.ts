/**
 * Advance-width contract of the bundled Bengali faces (fonts/README.md):
 * advances rewritten to Word's Nirmala UI (tools/build-bengali-font.py). A
 * regenerated woff2 that loses the normalization would drift Bengali line
 * breaks ~8% vs Word (Kohinoor Bangla fallback, prod corpus 20261021 sample 013).
 */
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { advanceEm, readWoff2 } from './helpers/woff2-metrics'

const FONTS = join(__dirname, '../src/renderer/fonts')
const regular = readWoff2(join(FONTS, 'GenOfficeBengali-Regular.woff2'))
const bold = readWoff2(join(FONTS, 'GenOfficeBengali-Bold.woff2'))

describe('GenOffice Bengali (Nirmala UI-normalized)', () => {
  it('space, danda and Bengali digits match Nirmala UI', () => {
    expect(advanceEm(regular, 0x20)).toBeCloseTo(0.274, 3)
    expect(advanceEm(regular, 0x0964)).toBeCloseTo(0.432, 3)
    expect(advanceEm(regular, 0x09e7)).toBeCloseTo(0.515, 3)
    expect(advanceEm(bold, 0x20)).toBeCloseTo(0.276, 3)
    expect(advanceEm(bold, 0x0985)).toBeCloseTo(1.099, 3)
  })

  it('Bengali letters and matras match Nirmala UI', () => {
    const expected: Record<number, number> = {
      0x0985: 0.986, // a
      0x0986: 1.208, // aa
      0x0995: 0.878, // ka
      0x09a4: 0.771, // ta
      0x09ae: 0.727, // ma
      0x09be: 0.251, // aa sign
      0x09bf: 0.25, // i sign
      0x09c7: 0.411, // e sign
    }
    for (const [cp, adv] of Object.entries(expected)) {
      expect(advanceEm(regular, Number(cp)), `U+${Number(cp).toString(16)}`).toBeCloseTo(adv, 3)
    }
  })
})
