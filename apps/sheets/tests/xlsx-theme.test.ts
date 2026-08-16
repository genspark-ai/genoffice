import { describe, expect, it } from 'vitest'

import { applyThemeState, ThemeStateError } from '../src/gateway/xlsx-theme'

const THEME =
  '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">' +
  '<a:themeElements><a:clrScheme name="Office">' +
  '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>' +
  '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="44546A"/></a:dk2>' +
  '<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
  '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>' +
  '<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
  '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>' +
  '<a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
  '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>' +
  '<a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
  '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>' +
  '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>' +
  '</a:clrScheme><a:fontScheme name="Office">' +
  '<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/></a:majorFont>' +
  '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/></a:minorFont>' +
  '</a:fontScheme></a:themeElements></a:theme>'

/// Theme index order: [lt1, dk1, lt2, dk2, accent1-6, hlink, folHlink].
const PALETTE = [
  '#FFFFFE',
  '#000001',
  '#EEEEEE',
  '#333333',
  '#111111',
  '#222222',
  '#444444',
  '#555555',
  '#666666',
  '#777777',
  '#888888',
  '#999999',
]

describe('applyThemeState', () => {
  it('rewrites clrScheme slots with the light/dark index swap', () => {
    const xml = applyThemeState(THEME, { colors: { name: 'Custom', values: PALETTE } })
    expect(xml).toContain('<a:clrScheme name="Custom">')
    expect(xml).toContain('<a:dk1><a:srgbClr val="000001"/></a:dk1>')
    expect(xml).toContain('<a:lt1><a:srgbClr val="FFFFFE"/></a:lt1>')
    expect(xml).toContain('<a:dk2><a:srgbClr val="333333"/></a:dk2>')
    expect(xml).toContain('<a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>')
    expect(xml).toContain('<a:accent1><a:srgbClr val="111111"/></a:accent1>')
    expect(xml).toContain('<a:folHlink><a:srgbClr val="999999"/></a:folHlink>')
    // fontScheme untouched
    expect(xml).toContain('typeface="Calibri Light"')
  })

  it('rewrites only the latin typefaces of the fontScheme', () => {
    const xml = applyThemeState(THEME, {
      fonts: { name: 'Georgia', major: 'Georgia', minor: 'Georgia' },
    })
    expect(xml).toContain('<a:fontScheme name="Georgia">')
    expect(xml).toContain('<a:majorFont><a:latin typeface="Georgia"/>')
    expect(xml).toContain('<a:minorFont><a:latin typeface="Georgia"/>')
    expect(xml).toContain('<a:ea typeface=""/>')
    // clrScheme untouched
    expect(xml).toContain('<a:accent1><a:srgbClr val="4472C4"/></a:accent1>')
  })

  it('rejects malformed palettes and themes without the target scheme', () => {
    expect(() =>
      applyThemeState(THEME, { colors: { name: 'X', values: PALETTE.slice(0, 11) } }),
    ).toThrow(ThemeStateError)
    expect(() =>
      applyThemeState(THEME, { colors: { name: 'X', values: [...PALETTE.slice(0, 11), 'red'] } }),
    ).toThrow(ThemeStateError)
    expect(() => applyThemeState('<a:theme/>', { colors: { name: 'X', values: PALETTE } })).toThrow(
      ThemeStateError,
    )
    expect(() =>
      applyThemeState('<a:theme/>', { fonts: { name: 'X', major: 'A', minor: 'B' } }),
    ).toThrow(ThemeStateError)
  })
})
