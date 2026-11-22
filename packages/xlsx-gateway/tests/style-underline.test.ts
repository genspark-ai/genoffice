import { describe, expect, it } from 'vitest'

import { StylesheetEditor } from '../src/gateway/xlsx-styles'

const STYLES = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>`

describe('whole-cell underline variants', () => {
  it('writes the double variant back and a bare <u/> for single', () => {
    const editor = new StylesheetEditor(STYLES)
    const double = editor.resolveStyle(0, { underline: true, underlineStyle: 'double' })
    const fonts = editor.serialize().match(/<font>[\s\S]*?<\/font>/g) ?? []
    expect(fonts.filter((font) => font.includes('<u val="double"/>'))).toHaveLength(1)

    // a plain underline (or an explicit single) writes the bare <u/>
    const single = editor.resolveStyle(0, { underline: true, underlineStyle: 'single' })
    const after = editor.serialize().match(/<font>[\s\S]*?<\/font>/g) ?? []
    expect(after.filter((font) => /<u\/>/.test(font))).toHaveLength(1)
    expect(single).not.toBe(double)
  })

  it('drops the underline (and its variant) when the delta turns it off', () => {
    const editor = new StylesheetEditor(STYLES)
    const base = editor.resolveStyle(0, { underline: true, underlineStyle: 'double' })
    const off = editor.resolveStyle(base, { underline: false })
    const fonts = editor.serialize().match(/<font>[\s\S]*?<\/font>/g) ?? []
    expect(fonts[off]).not.toContain('<u')
  })
})
