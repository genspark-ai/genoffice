import { describe, expect, it } from 'vitest'

import { StylesheetEditor } from '../src/gateway/xlsx-styles'

// A real Excel stylesheet: the default font carries a theme-provenance
// <scheme val="minor"/>, so the reader (and Excel) resolve the family from the
// theme and ignore the cached <name val>.
const STYLES =
  `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
  `<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>` +
  `<fills count="1"><fill><patternFill patternType="none"/></fill></fills>` +
  `<borders count="1"><border/></borders>` +
  `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>`

describe('font family on a scheme-derived font', () => {
  it('drops <scheme> when the family is set, so the choice is not reverted on reopen', () => {
    const editor = new StylesheetEditor(STYLES)
    const index = editor.resolveStyle(0, { fontFamily: 'Arial' })
    expect(index).toBe(1)
    const fonts = editor.serialize().match(/<font>[\s\S]*?<\/font>/g) ?? []
    expect(fonts[1]).toContain('<name val="Arial"/>')
    // The theme scheme would otherwise resolve the family back to Calibri.
    expect(fonts[1]).not.toContain('<scheme')
  })

  it('keeps other font attributes (size, family, color) on the derived font', () => {
    const editor = new StylesheetEditor(STYLES)
    editor.resolveStyle(0, { fontFamily: 'Arial' })
    const fonts = editor.serialize().match(/<font>[\s\S]*?<\/font>/g) ?? []
    expect(fonts[1]).toContain('<sz val="11"/>')
    expect(fonts[1]).toContain('<family val="2"/>')
    expect(fonts[1]).toContain('<color theme="1"/>')
  })

  it('leaves the base font untouched', () => {
    const editor = new StylesheetEditor(STYLES)
    editor.resolveStyle(0, { fontFamily: 'Arial' })
    const fonts = editor.serialize().match(/<font>[\s\S]*?<\/font>/g) ?? []
    expect(fonts[0]).toContain('<name val="Calibri"/>')
    expect(fonts[0]).toContain('<scheme val="minor"/>')
  })
})
