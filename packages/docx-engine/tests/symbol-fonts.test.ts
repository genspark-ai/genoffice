import { describe, expect, it } from 'vitest'
import {
  decodeSymbolChar,
  decodeSymbolText,
  isSymbolFont,
  parseDocx,
  toSymbolPua,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

describe('symbol font decoding helpers', () => {
  it('recognizes symbol-encoded fonts case-insensitively', () => {
    expect(isSymbolFont('Wingdings')).toBe(true)
    expect(isSymbolFont('symbol')).toBe(true)
    expect(isSymbolFont('Webdings')).toBe(true)
    expect(isSymbolFont('Calibri')).toBe(false)
    expect(isSymbolFont(null)).toBe(false)
  })

  it('maps raw bytes and their U+F0xx private-use forms alike', () => {
    expect(decodeSymbolChar('Wingdings', 0x6c)).toBe('●')
    expect(decodeSymbolChar('Wingdings', 0xf06c)).toBe('●')
    expect(decodeSymbolChar('Symbol', 0xf0b7)).toBe('•')
    expect(decodeSymbolChar('Wingdings', 0xf000)).toBeNull()
  })

  it('decodes whole strings only when every glyph maps', () => {
    expect(decodeSymbolText('Symbol', 'abg')).toBe('αβγ')
    expect(decodeSymbolText('Wingdings', 'l\u2013')).toBeNull()
    expect(decodeSymbolText('Calibri', 'abc')).toBeNull()
  })

  it('text runs keep pictograph glyphs in their font, take monochrome stand-ins', () => {
    expect(decodeSymbolText('Wingdings', 'C', { textGlyphsOnly: true })).toBeNull()
    expect(decodeSymbolText('Wingdings', '6', { textGlyphsOnly: true })).toBeNull()
    expect(decodeSymbolText('Wingdings', 'C')).toBe('\u{1F44D}')
    expect(decodeSymbolText('Wingdings', 'l\u00fc\u00f0', { textGlyphsOnly: true })).toBe('●✓→')
    expect(decodeSymbolText('Symbol', 'a \u00b7', { textGlyphsOnly: true })).toBe('α •')
    // ASCII approximations keep the font; ASCII that the font itself draws as ASCII decodes
    expect(decodeSymbolText('Webdings', '>', { textGlyphsOnly: true })).toBeNull()
    expect(decodeSymbolText('Wingdings', 'j', { textGlyphsOnly: true })).toBeNull()
    expect(decodeSymbolText('Webdings', '\uF03E', { textGlyphsOnly: true })).toBeNull()
    expect(decodeSymbolText('Symbol', '1+2=3', { textGlyphsOnly: true })).toBe('1+2=3')
  })

  it('the space glyph of every symbol font is a space', () => {
    expect(decodeSymbolChar('Symbol', 0xf020)).toBe(' ')
    expect(decodeSymbolChar('Webdings', 0x20)).toBe(' ')
  })

  it('maps the Wingdings 2 circle/square bullet series', () => {
    expect(decodeSymbolChar('Wingdings 2', 0x97)).toBe('●')
    expect(decodeSymbolChar('Wingdings 2', 0xf09b)).toBe('○')
    expect(decodeSymbolChar('Wingdings 2', 0xf09e)).toBe('◉')
    expect(decodeSymbolChar('Wingdings 2', 0xf0a1)).toBe('■')
    expect(decodeSymbolChar('Wingdings 2', 0xf0a4)).toBe('□')
  })

  it('maps the common Wingdings bullet, arrow and check-box glyphs', () => {
    expect(decodeSymbolChar('Wingdings', 0xf09f)).toBe('•')
    expect(decodeSymbolChar('Wingdings', 0xf09e)).toBe('·')
    expect(decodeSymbolChar('Wingdings', 0xf0a8)).toBe('□')
    expect(decodeSymbolChar('Wingdings', 0xf0d8)).toBe('➢')
    expect(decodeSymbolChar('Wingdings', 0xf0e0)).toBe('⇨')
    expect(decodeSymbolChar('Wingdings', 0xf0f0)).toBe('→')
    expect(decodeSymbolChar('Wingdings', 0xf0fb)).toBe('✗')
    expect(decodeSymbolChar('Wingdings', 0xf0fc)).toBe('✓')
    expect(decodeSymbolChar('Wingdings', 0xf0fd)).toBe('☒')
    expect(decodeSymbolChar('Wingdings', 0xf0fe)).toBe('☑')
    expect(decodeSymbolChar('Wingdings', 0xf0b7)).toBe('\u{1F550}')
  })

  it('maps the Wingdings 2 check marks and their boxed forms', () => {
    expect(decodeSymbolChar('Wingdings 2', 0xf04f)).toBe('✕')
    expect(decodeSymbolChar('Wingdings 2', 0xf050)).toBe('✓')
    expect(decodeSymbolChar('Wingdings 2', 0xf051)).toBe('☒')
    expect(decodeSymbolChar('Wingdings 2', 0xf052)).toBe('☑')
    expect(decodeSymbolChar('Wingdings 2', 0xf054)).toBe('☒')
  })

  it('maps Wingdings 3 triangles and Webdings solid shapes', () => {
    expect(decodeSymbolChar('Wingdings 3', 0xf075)).toBe('▶')
    expect(decodeSymbolChar('Wingdings 3', 0xf070)).toBe('▲')
    expect(decodeSymbolChar('Webdings', 0xf06e)).toBe('●')
    expect(decodeSymbolChar('Webdings', 0xf067)).toBe('■')
  })

  it('covers the whole glyph range of every symbol font, raw byte or U+F0xx alike', () => {
    const defined: Record<string, number[]> = {
      Symbol: [0x7f, 0xa0 - 1, 0xff],
      Wingdings: [0x7f, 0xff],
      'Wingdings 2': [0x7f, 0xfa, 0xfb, 0xfc, 0xfd, 0xfe, 0xff],
      'Wingdings 3': [
        0x7f, 0xf1, 0xf2, 0xf3, 0xf4, 0xf5, 0xf6, 0xf7, 0xf8, 0xf9, 0xfa, 0xfb, 0xfc, 0xfd, 0xfe,
        0xff,
      ],
      Webdings: [0x7f],
    }
    for (const [font, holes] of Object.entries(defined)) {
      for (let code = 0x21; code <= 0xff; code++) {
        const isHole = holes.includes(code) || (font === 'Symbol' && code >= 0x80 && code < 0xa0)
        const glyph = decodeSymbolChar(font, code)
        expect(glyph !== null, `${font} 0x${code.toString(16)}`).toBe(!isHole)
        expect(decodeSymbolChar(font, 0xf000 + code)).toBe(glyph)
      }
    }
  })

  it('maps the Symbol font: Greek, operators, arrows, card suits and the serif marks', () => {
    expect(decodeSymbolText('Symbol', 'abgdW')).toBe('αβγδΩ')
    expect(decodeSymbolChar('Symbol', 0x2d)).toBe('−')
    expect(decodeSymbolChar('Symbol', 0xf0ae)).toBe('→')
    expect(decodeSymbolChar('Symbol', 0xf0a8)).toBe('♦')
    expect(decodeSymbolChar('Symbol', 0xf0d3)).toBe('©')
    expect(decodeSymbolChar('Symbol', 0xf0e5)).toBe('∑')
    expect(decodeSymbolChar('Symbol', 0xf0f2)).toBe('∫')
  })

  it('maps the Wingdings pictographs, circled digits and arrows', () => {
    expect(decodeSymbolChar('Wingdings', 0x21)).toBe('✏')
    expect(decodeSymbolChar('Wingdings', 0x22)).toBe('✂')
    expect(decodeSymbolChar('Wingdings', 0x28)).toBe('☎')
    expect(decodeSymbolChar('Wingdings', 0x6f)).toBe('❑')
    expect(decodeSymbolChar('Wingdings', 0x8c)).toBe('❶')
    expect(decodeSymbolChar('Wingdings', 0x81)).toBe('①')
    expect(decodeSymbolChar('Wingdings', 0xab)).toBe('★')
    expect(decodeSymbolChar('Wingdings', 0xe8)).toBe('➡')
    expect(decodeSymbolChar('Wingdings', 0xfe)).toBe('☑')
    expect(decodeSymbolChar('Wingdings 3', 0x21)).toBe('←')
    expect(decodeSymbolChar('Wingdings 3', 0xf0a9)).toBe('⬅')
    expect(decodeSymbolChar('Webdings', 0x61)).toBe('✔')
  })

  it('never resolves to another private-use code point', () => {
    for (const font of ['Symbol', 'Wingdings', 'Wingdings 2', 'Wingdings 3', 'Webdings']) {
      for (let code = 0x21; code <= 0xff; code++) {
        const glyph = decodeSymbolChar(font, code)
        if (glyph === null) continue
        const cp = glyph.codePointAt(0) as number
        expect(cp >= 0xe000 && cp <= 0xf8ff, `${font} 0x${code.toString(16)}`).toBe(false)
      }
    }
  })

  it('normalizes raw glyph bytes to their U+F0xx form', () => {
    expect(toSymbolPua('l')).toBe('\uF06C')
    expect(toSymbolPua('·')).toBe('\uF0B7')
    expect(toSymbolPua('\uF0B7')).toBe('\uF0B7')
    expect(toSymbolPua(' l')).toBe(' \uF06C')
  })
})

describe('symbol runs in document.xml', () => {
  it('converts w:sym to the Unicode equivalent', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:t>勾:</w:t></w:r><w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r>' +
        '<w:r><w:sym w:font="Wingdings" w:char="6C"/></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].runs!.map((r) => r.text).join('')).toBe('勾:✓●')
  })

  it('renders a Wingdings 2 boxed X check box from w:sym', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:t>New</w:t></w:r><w:r><w:sym w:font="Wingdings 2" w:char="F054"/></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].runs!.map((r) => r.text).join('')).toBe('New☒')
  })

  it('keeps unknown w:sym glyphs as private-use characters', async () => {
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:sym w:font="Wingdings 2" w:char="F0FF"/></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].runs!.map((r) => r.text).join('')).toBe('\uF0FF')
  })

  it('decodes text runs carrying a symbol font and drops the font', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr><w:t>&#xF0B7;</w:t></w:r>' +
        '<w:r><w:t xml:space="preserve"> 后文</w:t></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    const run = doc.blocks[0].runs![0]
    expect(doc.blocks[0].runs!.map((r) => r.text).join('')).toBe('• 后文')
    expect(run.font).toBeUndefined()
    expect(run.rawRPr ?? '').not.toContain('w:rFonts')
  })

  it('decodes a whole Symbol-font text run to Greek', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr><w:t>a + b = g</w:t></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    const run = doc.blocks[0].runs![0]
    expect(run.text).toBe('α + β = γ')
    expect(run.font).toBeUndefined()
  })

  it('decodes by the Latin slot when an eastAsia twin names a text face', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:rPr><w:rFonts w:ascii="Symbol" w:eastAsia="Times New Roman" w:hAnsi="Symbol" w:cs="Times New Roman"/></w:rPr><w:t>&#xF021;&#xF022;</w:t></w:r>' +
        '<w:r><w:rPr><w:rFonts w:eastAsia="Times New Roman" w:hAnsi="Symbol"/></w:rPr><w:t>&#xF024;</w:t></w:r>' +
        '<w:r><w:rPr><w:rFonts w:ascii="Wingdings" w:eastAsia="Times New Roman" w:hAnsi="Wingdings"/></w:rPr><w:t>&#xF043;</w:t></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    const runs = doc.blocks[0].runs!
    expect(runs.map((r) => r.text)).toEqual(['!\u2200', '\u2203', '\uF043'])
    expect(runs[0].font).toBeUndefined()
    expect(runs[0].fontAscii).toBeUndefined()
    expect(runs[1].font).toBeUndefined()
    expect(runs[2].fontAscii).toBe('Wingdings')
  })

  it('leaves undecodable symbol-font runs untouched', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr><w:t>l\u2013</w:t></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    const run = doc.blocks[0].runs![0]
    expect(run.text).toBe('l\u2013')
    expect(run.font).toBe('Wingdings')
  })

  it('keeps a pictograph text run in its symbol font', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:rPr><w:rFonts w:ascii="Webdings" w:hAnsi="Webdings"/></w:rPr><w:t>b</w:t></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    const run = doc.blocks[0].runs![0]
    expect(run.text).toBe('b')
    expect(run.font).toBe('Webdings')
  })

  it('decodes numeric character references in w:t', async () => {
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>&#x2713;&#65;</w:t></w:r></w:p>',
    })
    const doc = await parseDocx(bytes)
    expect(doc.blocks[0].runs!.map((r) => r.text).join('')).toBe('✓A')
  })
})

describe('numbering level marker font', () => {
  it('parses w:rPr/w:rFonts of a bullet level', async () => {
    const numberingXml =
      XML_DECL +
      '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:abstractNum w:abstractNumId="0">' +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="l"/>' +
      '<w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr></w:lvl>' +
      '</w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
      '</w:numbering>'
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>项</w:t></w:r></w:p>',
      numberingXml,
    })
    const doc = await parseDocx(bytes)
    expect(doc.numbering.get('1')!.levels[0].font).toBe('Wingdings')
  })
})
