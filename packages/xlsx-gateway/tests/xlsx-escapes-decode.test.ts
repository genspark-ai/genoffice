import { describe, expect, it } from 'vitest'
import { decodeXlsxEscapes, decodeXmlText, encodeXlsxEscapes } from '../src/gateway/xlsx-escapes'

describe('decodeXmlText numeric character references', () => {
  it('decodes in-range references and named entities', () => {
    expect(decodeXmlText('&#65;&#x4e2d;&amp;&quot;')).toBe('A中&"')
  })

  it('keeps out-of-range references as text', () => {
    expect(decodeXmlText('a&#x110000;b')).toBe('a&#x110000;b')
  })

  it('keeps lone surrogate references as text so a resave stays valid XML', () => {
    expect(decodeXmlText('&#xD800;')).toBe('&#xD800;')
    const text = decodeXmlText('x&#xD83D;DE00y')
    expect(text).not.toMatch(/[\uD800-\uDFFF]/)
    expect(encodeXlsxEscapes(text)).not.toMatch(/[\uD800-\uDFFF]/)
  })

  it('still rejects surrogate _xHHHH_ escapes', () => {
    expect(decodeXlsxEscapes('a_xD800_b')).toBe('a_xD800_b')
    expect(decodeXlsxEscapes('a_xD83D_DE00b')).toBe('a_xD83D_DE00b')
  })
})
