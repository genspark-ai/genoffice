import { describe, expect, it } from 'vitest'
import { parseFileToText } from '../src/index'
import { resolveTarget } from '../src/opc'
import { writeFixture } from './helpers/fixtures'

describe('parseFileToText: plain-text formats', () => {
  const cases: Array<[string, string]> = [
    ['sample.txt', 'plain text hello'],
    ['sample.md', '# Title\n\nBody paragraph'],
    ['sample.csv', 'a,b,c\n1,2,3'],
    ['sample.tsv', 'a\tb\tc\n1\t2\t3'],
    ['sample.json', '{"key":"value"}'],
    ['sample.xml', '<root><item>value</item></root>'],
    ['sample.html', '<html><body><p>page</p></body></html>'],
    ['sample.py', 'def hello():\n    return "world"'],
  ]

  for (const [name, content] of cases) {
    it(`reads ${name} verbatim`, async () => {
      const path = writeFixture(name, content)
      const result = await parseFileToText(path)
      expect(result).toEqual({ ok: true, kind: 'text', text: content })
    })
  }

  it('is case-insensitive on the extension', async () => {
    const path = writeFixture('UPPER.TXT', 'upper')
    const result = await parseFileToText(path)
    expect(result.ok).toBe(true)
    expect(result.text).toBe('upper')
  })

  it('strips a UTF-8 BOM instead of leaking it into the text', async () => {
    const bytes = Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from('bom hello', 'utf8')])
    const result = await parseFileToText(writeFixture('bom.txt', bytes))
    expect(result).toEqual({ ok: true, kind: 'text', text: 'bom hello' })
  })

  it('decodes UTF-16LE and UTF-16BE BOM files instead of returning mojibake', async () => {
    const le = writeFixture(
      'utf16le.txt',
      Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from('héllo le', 'utf16le')]),
    )
    expect(await parseFileToText(le)).toEqual({ ok: true, kind: 'text', text: 'héllo le' })
    const beBody = Buffer.from('héllo be', 'utf16le')
    beBody.swap16()
    const be = writeFixture('utf16be.txt', Buffer.concat([Buffer.from([0xfe, 0xff]), beBody]))
    expect(await parseFileToText(be)).toEqual({ ok: true, kind: 'text', text: 'héllo be' })
  })

  it('rejects only a BOM-declared UTF-32 file, whose encoding cannot be decoded', async () => {
    const utf32 = writeFixture(
      'utf32le.txt',
      Buffer.from([0xff, 0xfe, 0x00, 0x00, 0x41, 0x00, 0x00, 0x00]),
    )
    const utf32Result = await parseFileToText(utf32)
    expect(utf32Result.ok).toBe(false)
    expect(utf32Result.error).toContain('encoding')
  })

  it('falls back to windows-1252 for latin-1 bytes instead of dropping the file', async () => {
    const latin1 = writeFixture('latin1.csv', Buffer.from('caf\xe9,1', 'latin1'))
    const result = await parseFileToText(latin1)
    expect(result.ok).toBe(true)
    expect(result.kind).toBe('text')
    expect(result.text).toBe('café,1')
  })

  it('returns the whole text for valid UTF-8 with a single stray 0xFF byte', async () => {
    const bytes = Buffer.concat([
      Buffer.from('hello '.repeat(500), 'utf8'),
      Buffer.from([0xff]),
      Buffer.from(' world'.repeat(500), 'utf8'),
    ])
    const result = await parseFileToText(writeFixture('stray-ff.txt', bytes))
    expect(result.ok).toBe(true)
    expect(result.text).toContain('hello hello')
    expect(result.text).toContain('world world')
    // windows-1252 maps every byte to a character, so nothing is truncated
    expect(result.text?.length).toBe(bytes.length)
  })

  it('never rejects valid UTF-8 that legitimately contains U+FFFD', async () => {
    const result = await parseFileToText(writeFixture('fffd.txt', 'bad \uFFFD char'))
    expect(result).toEqual({ ok: true, kind: 'text', text: 'bad \uFFFD char' })
  })

  it('keeps GBK / Shift-JIS text indexed (mojibake) instead of erroring it away', async () => {
    // "中文" in GBK; "日本語" in Shift-JIS — neither is valid UTF-8
    const gbk = writeFixture('gbk.txt', Buffer.from([0xd6, 0xd0, 0xce, 0xc4]))
    const gbkResult = await parseFileToText(gbk)
    expect(gbkResult.ok).toBe(true)
    expect(gbkResult.kind).toBe('text')
    expect(gbkResult.text?.length).toBe(4)

    const shiftJis = writeFixture('shiftjis.txt', Buffer.from([0x93, 0xfa, 0x96, 0x7b, 0x8c, 0xea]))
    const shiftJisResult = await parseFileToText(shiftJis)
    expect(shiftJisResult.ok).toBe(true)
    expect(shiftJisResult.text?.length).toBe(6)
  })

  it('fails gracefully on a missing file', async () => {
    const result = await parseFileToText('/nonexistent/nowhere.txt')
    expect(result.ok).toBe(false)
    expect(result.error).toBeTruthy()
  })
})

describe('OPC relationship targets', () => {
  it('decodes percent-encoded targets and normalizes backslashes', () => {
    expect(resolveTarget('ppt/presentation.xml', 'slides/slide%201.xml')).toBe(
      'ppt/slides/slide 1.xml',
    )
    expect(resolveTarget('ppt/slides/slide1.xml', '..\\media\\image%201.png')).toBe(
      'ppt/media/image 1.png',
    )
  })
})

describe('parseFileToText: images and unsupported', () => {
  it('flags png as image without extracting text', async () => {
    const path = writeFixture('pic.png', Buffer.from('89504e47', 'hex'))
    const result = await parseFileToText(path)
    expect(result).toEqual({ ok: true, kind: 'image', mime: 'image/png' })
  })

  it.each([
    ['pic.jpg', 'image/jpeg'],
    ['pic.jpeg', 'image/jpeg'],
    ['pic.gif', 'image/gif'],
    ['pic.webp', 'image/webp'],
  ])('maps %s to mime %s', async (name, mime) => {
    const path = writeFixture(name, Buffer.from([0]))
    const result = await parseFileToText(path)
    expect(result.kind).toBe('image')
    expect(result.mime).toBe(mime)
  })

  it('rejects unknown extensions as unsupported', async () => {
    const path = writeFixture('archive.zip', Buffer.from([0x50, 0x4b]))
    const result = await parseFileToText(path)
    expect(result.ok).toBe(false)
    expect(result.kind).toBe('unsupported')
    expect(result.error).toContain('.zip')
  })
})
