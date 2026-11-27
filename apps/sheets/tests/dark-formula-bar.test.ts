import { describe, expect, it } from 'vitest'
import { compressCssColor, compressRgb, SHEETS_DARK_FLOOR } from '../src/renderer/dark-canvas-color'

/**
 * Univer's dark-mode inversion maps authored white to #000. The formula bar's
 * page is white, so without compression it becomes a black hole beside the
 * #1e1e1e chrome. These are the exact colors the canvas color service sees.
 */
describe('dark formula-bar canvas colors', () => {
  it('compresses inverted white onto the chrome floor and keeps black white', () => {
    expect(compressRgb([0, 0, 0])).toEqual([30, 30, 30])
    expect(compressRgb([255, 255, 255])).toEqual([255, 255, 255])
    expect(SHEETS_DARK_FLOOR).toEqual([30, 30, 30])
  })

  it('compresses hex colors while keeping alpha', () => {
    expect(compressCssColor('#000000')).toBe('#1e1e1e')
    expect(compressCssColor('#ffffff')).toBe('#ffffff')
    expect(compressCssColor('#808080')).toBe('#8f8f8f')
    expect(compressCssColor('#00000080')).toBe('#1e1e1e80')
    expect(compressCssColor('#0008')).toBe('#1e1e1e8')
  })

  it('compresses rgb/rgba colors while keeping alpha', () => {
    expect(compressCssColor('rgb(0, 0, 0)')).toBe('rgb(30,30,30)')
    expect(compressCssColor('rgba(0,0,0,0.5)')).toBe('rgba(30,30,30, 0.5)')
    expect(compressCssColor('rgba(255, 255, 255, 1)')).toBe('rgba(255,255,255, 1)')
  })

  it('keeps hues in their direction', () => {
    const [r, g, b] = compressRgb([200, 20, 40])
    expect(r).toBeGreaterThan(g)
    expect(r).toBeGreaterThan(b)
    expect(g).toBeGreaterThan(30)
  })

  it('leaves colors it does not own alone', () => {
    expect(compressCssColor('transparent')).toBe('transparent')
    expect(compressCssColor('@themePrimary')).toBe('@themePrimary')
  })
})
