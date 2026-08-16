import { describe, expect, it } from 'vitest'

import type { WorkbookCellStyle } from '../src/shared/desktop-api'
import { applyTint, COLOR_SCHEMES, rethemeStyles, THEME_PRESETS } from '../src/renderer/themes'

describe('theme presets', () => {
  it('store 12-slot palettes in theme index order (lt1 first)', () => {
    const office = COLOR_SCHEMES.find((scheme) => scheme.id === 'office')!
    expect(office.values).toHaveLength(12)
    expect(office.values[0]).toBe('#FFFFFF') // lt1
    expect(office.values[1]).toBe('#000000') // dk1
    expect(office.values[4]).toBe('#4472C4') // accent1
    expect(office.values[10]).toBe('#0563C1') // hlink
    for (const preset of THEME_PRESETS) {
      expect(preset.colors.values).toHaveLength(12)
      expect(preset.fonts.major.length).toBeGreaterThan(0)
    }
  })
})

describe('applyTint', () => {
  it('matches the sidecar transform at the extremes', () => {
    expect(applyTint('#4472C4', 0)).toBe('#4472C4')
    expect(applyTint('#4472C4', 1)).toBe('#FFFFFF')
    expect(applyTint('#4472C4', -1)).toBe('#000000')
  })

  it('lightens on positive tint and darkens on negative, keeping hue', () => {
    const lighter = applyTint('#4472C4', 0.6)
    const darker = applyTint('#4472C4', -0.25)
    const luminance = (hex: string): number =>
      parseInt(hex.slice(1, 3), 16) + parseInt(hex.slice(3, 5), 16) + parseInt(hex.slice(5), 16)
    expect(luminance(lighter)).toBeGreaterThan(luminance('#4472C4'))
    expect(luminance(darker)).toBeLessThan(luminance('#4472C4'))
    // Excel's accent1 60%-lighter swatch for the Office palette.
    expect(lighter).toBe('#B4C7E7')
  })
})

const baseStyle: WorkbookCellStyle = {
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  wrapText: false,
  diagonalUp: false,
  diagonalDown: false,
}

describe('rethemeStyles', () => {
  const midnight = COLOR_SCHEMES.find((scheme) => scheme.id === 'midnight')!

  it('re-resolves provenance colors and theme-scheme fonts', () => {
    const styles: WorkbookCellStyle[] = [
      {
        ...baseStyle,
        fontColor: '#4472C4',
        fontColorTheme: 4,
        fontFamily: 'Calibri',
        fontScheme: 'minor',
      },
      { ...baseStyle, fillColor: '#B4C7E7', fillColorTheme: 4, fillColorTint: 0.6 },
    ]
    const next = rethemeStyles(styles, midnight.values, { major: 'Georgia', minor: 'Georgia' })
    expect(next[0]!.fontColor).toBe('#4A9EDE')
    expect(next[0]!.fontFamily).toBe('Georgia')
    expect(next[1]!.fillColor).toBe(applyTint('#4A9EDE', 0.6))
  })

  it('leaves literal-color styles untouched by reference', () => {
    const styles: WorkbookCellStyle[] = [
      { ...baseStyle, fontColor: '#FF0000', fillColor: '#00FF00', fontFamily: 'Arial' },
    ]
    const next = rethemeStyles(styles, midnight.values, { major: 'Georgia', minor: 'Georgia' })
    expect(next[0]).toBe(styles[0])
  })
})
