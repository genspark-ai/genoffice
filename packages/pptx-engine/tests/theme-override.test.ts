/**
 * Chart-part theme overrides (<a:themeOverride>): the chart's schemeClr refs
 * resolve against ITS clrScheme, not the deck theme's — decks legitimately
 * remap accents per chart (prod_050: accent1 and accent5 swapped, so every
 * series painted the wrong blue without this).
 */
import { describe, it, expect } from 'vitest'
import { parseTheme, themeWithOverride } from '../src/theme'

const BASE = parseTheme(
  `<a:theme><a:themeElements><a:clrScheme name="b">` +
    `<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>` +
    `<a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>` +
    `<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>` +
    `<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>` +
    `<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>` +
    `<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>` +
    `</a:clrScheme><a:fontScheme name="f"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>` +
    `<a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`,
)

const OVERRIDE =
  `<a:themeOverride xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
  `<a:clrScheme name="o">` +
  `<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>` +
  `<a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="DDDDDD"/></a:lt2>` +
  `<a:accent1><a:srgbClr val="5B9BD5"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>` +
  `<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>` +
  `<a:accent5><a:srgbClr val="4472C4"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>` +
  `<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>` +
  `</a:clrScheme></a:themeOverride>`

describe('themeWithOverride', () => {
  it('override clrScheme wins; base fonts survive when the override has none', () => {
    const t = themeWithOverride({ ...BASE, clrMap: { bg1: 'dk1' } }, OVERRIDE)
    expect(t.colors.accent5).toBe('#4472C4') // swapped relative to the base
    expect(t.colors.accent1).toBe('#5B9BD5')
    expect(t.minorFont).toBe('Calibri') // fontScheme absent from the override → base kept
    expect(t.clrMap).toEqual({ bg1: 'dk1' }) // clrMap is a master concern, never overridden
  })

  it('an override fontScheme replaces the base fonts', () => {
    const ov = OVERRIDE.replace(
      '</a:clrScheme></a:themeOverride>',
      `</a:clrScheme><a:fontScheme name="of"><a:majorFont><a:latin typeface="Georgia"/></a:majorFont>` +
        `<a:minorFont><a:latin typeface="Verdana"/></a:minorFont></a:fontScheme></a:themeOverride>`,
    )
    const t = themeWithOverride(BASE, ov)
    expect(t.minorFont).toBe('Verdana')
    expect(t.majorFont).toBe('Georgia')
  })

  it('a malformed override leaves the base theme untouched', () => {
    expect(themeWithOverride(BASE, '<a:nope/>')).toEqual(BASE)
  })
})
