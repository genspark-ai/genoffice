import { describe, expect, it } from 'vitest'

import { StylesheetEditor } from '../src/gateway/xlsx-styles'

const STYLES = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>`

describe('whole-cell vertAlign', () => {
  it('writes <vertAlign> into a derived font and null removes it', () => {
    const editor = new StylesheetEditor(STYLES)
    const superscript = editor.resolveStyle(0, { vertAlign: 'superscript' })
    expect(superscript).toBe(1)
    const xml = editor.serialize()
    const fonts = xml.match(/<font>[\s\S]*?<\/font>/g) ?? []
    expect(fonts[1]).toContain('<vertAlign val="superscript"/>')
    expect(fonts[1]).toContain('<name val="Calibri"/>')

    const subscript = editor.resolveStyle(superscript, { vertAlign: 'subscript' })
    const baseline = editor.resolveStyle(superscript, { vertAlign: null })
    const after = editor.serialize().match(/<font>[\s\S]*?<\/font>/g) ?? []
    expect(after.filter((font) => font.includes('<vertAlign val="subscript"/>'))).toHaveLength(1)
    expect(subscript).not.toBe(baseline)
    expect(after[0]).not.toContain('vertAlign')
    expect(after.filter((font) => font.includes('vertAlign'))).toHaveLength(2)
  })
})
