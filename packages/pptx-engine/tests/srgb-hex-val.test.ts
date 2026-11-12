/**
 * `a:srgbClr/@val` is six hex digits. The three readers that built `#` +
 * `String(val)` produced the truthy string `#UNDEFINED` for a missing value,
 * and hexToRgb parses 'UN'/'DE'/'FI' leniently — so a broken attribute became
 * a bright green fill (`#00DE0F`) instead of no colour at all.
 */
import { describe, it, expect } from 'vitest'
import { resolveColorNode } from '../src/color'
import { resolveTableStyle } from '../src/table-style'
import { parseTheme, resolveSchemeColor, type Theme } from '../src/theme'
import { srgbHexValue } from '../src/xml-utils'

const THEME_XML =
  '<?xml version="1.0"?><a:theme xmlns:a="a"><a:themeElements><a:clrScheme name="x">' +
  '<a:dk1><a:srgbClr/></a:dk1>' +
  '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="44546A"/></a:dk2>' +
  '<a:accent1><a:srgbClr val="RED"/></a:accent1>' +
  '<a:accent2><a:srgbClr val="4472C4"/></a:accent2>' +
  '</a:clrScheme></a:themeElements></a:theme>'

describe('srgbClr values that are not six hex digits', () => {
  it('accepts a well-formed value, with or without an alpha pair', () => {
    expect(srgbHexValue('4472c4')).toBe('#4472C4')
    expect(srgbHexValue('FF000080')).toBe('#FF000080')
  })

  it('refuses a missing, empty or malformed value', () => {
    expect(srgbHexValue(undefined)).toBeUndefined()
    expect(srgbHexValue('')).toBeUndefined()
    expect(srgbHexValue('UNDEFINED')).toBeUndefined()
    expect(srgbHexValue('RED')).toBeUndefined()
    expect(srgbHexValue('4472C')).toBeUndefined()
  })

  it('leaves the run/fill colour unresolved instead of inventing green', () => {
    expect(resolveColorNode({ 'a:srgbClr': {} }, undefined)).toBeUndefined()
    expect(resolveColorNode({ 'a:srgbClr': { '@_val': 'RED' } }, undefined)).toBeUndefined()
  })

  it('leaves a theme slot out of the palette', () => {
    const theme = parseTheme(THEME_XML)
    expect(theme.colors.dk1).toBeUndefined()
    expect(theme.colors.accent1).toBeUndefined()
    expect(theme.colors.dk2).toBe('#44546A')
    expect(theme.colors.accent2).toBe('#4472C4')
    // the system-colour branch is untouched
    expect(resolveSchemeColor('lt1', theme)).toBe('#FFFFFF')
  })

  it('leaves a table-style fill unset', () => {
    const style = (id: string, value: string) =>
      `<a:tblStyle styleId="{${id}}" styleName="${id}"><a:wholeTbl><a:tcStyle><a:fill>` +
      `<a:solidFill><a:srgbClr ${value}/></a:solidFill>` +
      '</a:fill></a:tcStyle></a:wholeTbl></a:tblStyle>'
    const xml =
      '<?xml version="1.0"?><a:tblStyleLst xmlns:a="a" def="{B}">' +
      style('B', 'val="UNDEFINED"') +
      style('C', 'val=""') +
      '</a:tblStyleLst>'
    const theme = { colors: {} } as Theme
    expect(resolveTableStyle('{B}', xml, theme)?.wholeTbl?.fill).toBeUndefined()
    expect(resolveTableStyle('{C}', xml, theme)?.wholeTbl?.fill).toBeUndefined()
  })
})
