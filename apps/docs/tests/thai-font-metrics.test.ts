/**
 * Advance-width contract of the bundled Thai faces (fonts/README.md): Noto Sans
 * Thai outlines with advances and vertical metrics rewritten to the Word faces
 * (tools/build-thai-font.py). Word for Mac lays TH SarabunPSK out real and
 * substitutes Cordia New for every other Thai family (probe 2026-11-03); a
 * regenerated woff2 that loses the normalization drifts Thai line breaks ~1.75x
 * (Chromium's Thonburi fallback). The outlines must be fitted too: a build
 * that only rewrote advances drew 1em Noto letters over 0.4em advances, so
 * every cluster piled onto the previous one.
 */
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { cssFontFamily, lineHeightFactor } from '../src/renderer/text-layout/font-rules'
import { resolveDeclaredFace } from '../src/renderer/font-resolve'
import { advanceEm, glyphBoundsEm, readWoff2 } from './helpers/woff2-metrics'

const FONTS = join(__dirname, '../src/renderer/fonts')
const cordia = readWoff2(join(FONTS, 'GenOfficeCordia-Regular.woff2'))
const cordiaBold = readWoff2(join(FONTS, 'GenOfficeCordia-Bold.woff2'))
const sarabun = readWoff2(join(FONTS, 'GenOfficeSarabun-Regular.woff2'))
const sarabunBold = readWoff2(join(FONTS, 'GenOfficeSarabun-Bold.woff2'))

describe('GenOffice Cordia / Sarabun (Word Thai metrics)', () => {
  it('Thai letters and marks match Cordia New', () => {
    const expected: Record<number, number> = {
      0x0e01: 0.4, // ko kai
      0x0e23: 0.3011, // ro rua
      0x0e32: 0.3468, // sara aa
      0x0e35: 0, // sara ii (combining)
      0x0e40: 0.2019, // sara e
      0x0020: 0.1822,
    }
    for (const [cp, adv] of Object.entries(expected)) {
      expect(advanceEm(cordia, Number(cp)), `U+${Number(cp).toString(16)}`).toBeCloseTo(adv, 3)
    }
  })

  it('Thai letters match TH SarabunPSK', () => {
    expect(advanceEm(sarabun, 0x0e01)).toBeCloseTo(0.3859, 3)
    expect(advanceEm(sarabun, 0x0e32)).toBeCloseTo(0.3161, 3)
    expect(advanceEm(sarabun, 0x0020)).toBeCloseTo(0.2161, 3)
  })

  it('bold faces carry the bold Word hmtx', () => {
    expect(advanceEm(cordiaBold, 0x0e01)).toBeCloseTo(0.4409, 3)
    expect(advanceEm(cordiaBold, 0x0e35)).toBe(0)
    expect(advanceEm(sarabunBold, 0x0e01)).toBeCloseTo(0.391, 3)
    expect(advanceEm(sarabunBold, 0x0e32)).toBeCloseTo(0.336, 3)
  })

  it('outlines fit the Word design size: ink inside the advance, consonant top ~0.4em', () => {
    for (const face of [cordia, cordiaBold, sarabun, sarabunBold]) {
      for (const cp of [0x0e01, 0x0e07, 0x0e19, 0x0e23]) {
        const [xMin, , xMax, yMax] = glyphBoundsEm(face, cp)!
        expect(xMin, `U+${cp.toString(16)} xMin`).toBeGreaterThanOrEqual(-0.01)
        expect(xMax, `U+${cp.toString(16)} xMax`).toBeLessThanOrEqual(advanceEm(face, cp) + 0.01)
        expect(yMax, `U+${cp.toString(16)} yMax`).toBeGreaterThan(0.38)
        expect(yMax, `U+${cp.toString(16)} yMax`).toBeLessThan(0.43)
      }
      const [, yMin] = glyphBoundsEm(face, 0x0e48)!
      expect(yMin, 'mai ek sits above the consonant top').toBeGreaterThan(0.43)
    }
  })

  it('missing Thai families chain to the Cordia face, TH SarabunPSK to its own', () => {
    expect(cssFontFamily('TH SarabunIT\u0e59')).toMatch(
      /^['"]TH SarabunIT\u0e59['"],'GenOffice Cordia'/,
    )
    expect(cssFontFamily('Angsana New')).toContain("'GenOffice Cordia'")
    expect(cssFontFamily('TH SarabunPSK')).toContain("'GenOffice Sarabun'")
    expect(cssFontFamily('Tahoma')).not.toContain('GenOffice Cordia')
  })

  it('line factors follow the Word faces (Cordia hhea 1.147, Sarabun 1.31; shipped faces real)', () => {
    expect(resolveDeclaredFace('TH SarabunIT\u0e59')?.family).toBe('Cordia New')
    expect(lineHeightFactor('TH SarabunIT\u0e59')).toBeCloseTo(1.147, 3)
    expect(lineHeightFactor('Cordia New')).toBeCloseTo(1.147, 3)
    expect(lineHeightFactor('TH SarabunPSK')).toBeCloseTo(1.31, 3)
    expect(lineHeightFactor('Angsana New')).toBe(1.3508)
    expect(lineHeightFactor('Browallia New')).toBe(1.2495)
  })
})
