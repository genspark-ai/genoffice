import { describe, it, expect } from 'vitest'
import { escapeXmlText, escapeXmlAttr, hasDefaultFor } from '../src/xml-utils'

describe('escapeXmlText XML 1.0 character defence', () => {
  it('escapes markup characters', () => {
    expect(escapeXmlText('a<b>&c')).toBe('a&lt;b&gt;&amp;c')
  })

  it('degrades VT/FF (PDF-extracted text) to a space', () => {
    expect(escapeXmlText(`para${String.fromCharCode(0x0b)}break`)).toBe('para break')
    expect(escapeXmlText(`page${String.fromCharCode(0x0c)}feed`)).toBe('page feed')
  })

  it('drops other forbidden chars: C0 controls, U+FFFE/FFFF, lone surrogates', () => {
    expect(escapeXmlText(`a${String.fromCharCode(0)}b${String.fromCharCode(8)}c`)).toBe('abc')
    expect(escapeXmlText('x￾y￿z')).toBe('xyz')
    expect(escapeXmlText('a\uD800b')).toBe('ab')
  })

  it('keeps legal whitespace and paired surrogates', () => {
    expect(escapeXmlText('a\tb\nc\rd')).toBe('a\tb\nc\rd')
    expect(escapeXmlText('emoji \u{1F600}!')).toBe('emoji \u{1F600}!')
  })

  it('escapeXmlAttr inherits the defence', () => {
    expect(escapeXmlAttr(`n${String.fromCharCode(1)}"q"`)).toBe('n&quot;q&quot;')
  })
})

describe('hasDefaultFor', () => {
  it('compares the extension literally, including regex metacharacters', () => {
    expect(hasDefaultFor('<Types><Default Extension="a+b"/></Types>', 'a+b')).toBe(true)
    expect(hasDefaultFor('<Types><Default Extension="ab"/></Types>', 'a+b')).toBe(false)
    expect(hasDefaultFor("<Types><Default Extension='png'/></Types>", 'png')).toBe(true)
    expect(hasDefaultFor('<Types><Default ContentType="x" Extension="jpg"/></Types>', 'jpg')).toBe(
      true,
    )
    expect(hasDefaultFor('<Types><Override PartName="/a"/></Types>', 'jpg')).toBe(false)
  })
})
