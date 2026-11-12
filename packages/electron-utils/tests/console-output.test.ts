import { describe, expect, it } from 'vitest'
import {
  decodeChunks,
  decodeChunksUtf8,
  decodeCommandOutput,
  decodeUtf8,
} from '../src/console-output'

// GBK bytes for U+5DE5 U+4F5C U+8868 ("worksheet"), as emitted by reg.exe on a
// zh-CN Windows install.
const gbkName = Buffer.concat([
  Buffer.from('Microsoft Excel ', 'ascii'),
  Buffer.from([0xb9, 0xa4, 0xd7, 0xf7, 0xb1, 0xed]),
])

// Shift-JIS (CP932) bytes for U+30EF U+30FC U+30AF U+30B7 U+30FC U+30C8.
const sjisName = Buffer.from([
  0x83, 0x8f, 0x81, 0x5b, 0x83, 0x4e, 0x83, 0x56, 0x81, 0x5b, 0x83, 0x67,
])

describe('decodeCommandOutput', () => {
  it('recovers the configured OEM code page instead of replacement boxes', () => {
    expect(gbkName.toString('utf8')).toContain('\uFFFD')
    expect(decodeCommandOutput(gbkName, 'win32', { codePage: 936 })).toBe(
      'Microsoft Excel \u5de5\u4f5c\u8868',
    )
  })

  it('decodes a non-GBK OEM page with its own mapping', () => {
    expect(decodeCommandOutput(sjisName, 'win32', { codePage: 932 })).toBe(
      '\u30ef\u30fc\u30af\u30b7\u30fc\u30c8',
    )
  })

  it('returns replacement characters for an unknown code page', () => {
    const out = decodeCommandOutput(Buffer.from('f\xfcr B\xfcro', 'latin1'), 'win32', {
      codePage: 99999,
    })
    expect(out).toContain('\uFFFD')
  })

  it('does not guess at an OEM page that has no decoder', () => {
    const out = decodeCommandOutput(Buffer.from([0x66, 0x81, 0x72]), 'win32', { codePage: 850 })
    expect(out).toContain('\uFFFD')
  })

  it('leaves ASCII output untouched on Windows', () => {
    const ascii = Buffer.from('    (Default)    REG_SZ    Word.Document.12\r\n', 'ascii')
    expect(decodeCommandOutput(ascii, 'win32', { codePage: 936 })).toBe(ascii.toString('utf8'))
  })

  it('keeps genuine UTF-8 output on Windows', () => {
    const utf8 = Buffer.from('Microsoft PowerPoint \u6f14\u793a\u6587\u7a3f', 'utf8')
    expect(decodeCommandOutput(utf8, 'win32', { codePage: 936 })).toBe(
      'Microsoft PowerPoint \u6f14\u793a\u6587\u7a3f',
    )
  })

  it('honours UTF-16 BOMs emitted by redirected PowerShell', () => {
    const text = '\u8bc1\u4e66 Subject'
    const le = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(text, 'utf16le')])
    expect(decodeCommandOutput(le, 'win32')).toBe(text)
    const be = Buffer.concat([Buffer.from([0xfe, 0xff]), Buffer.from(text, 'utf16le').swap16()])
    expect(decodeCommandOutput(be, 'win32')).toBe(text)
  })

  it('uses UTF-8 on macOS and Linux', () => {
    const utf8 = Buffer.from('WPS Office \u2014 \u8868\u683c', 'utf8')
    expect(decodeCommandOutput(utf8, 'linux')).toBe('WPS Office \u2014 \u8868\u683c')
    expect(decodeCommandOutput(utf8, 'darwin')).toBe('WPS Office \u2014 \u8868\u683c')
  })

  it('passes through an already-decoded string', () => {
    expect(decodeCommandOutput('Word Document', 'win32')).toBe('Word Document')
  })
})

describe('decodeChunks', () => {
  it('decodes a multi-byte character split across two chunks', () => {
    const bytes = Buffer.from('\u5de5\u4f5c\u8868', 'utf8')
    const cut = 2
    const chunks = [bytes.subarray(0, cut), bytes.subarray(cut)]
    expect(chunks.map((c) => c.toString()).join('')).toContain('\uFFFD')
    expect(decodeChunks(chunks)).toBe('\u5de5\u4f5c\u8868')
  })

  it('joins string chunks without re-encoding them', () => {
    expect(decodeChunks(['Microsoft ', 'Excel \u5de5\u4f5c\u8868'])).toBe(
      'Microsoft Excel \u5de5\u4f5c\u8868',
    )
  })

  it('decodes a Windows OEM stream as one unit', () => {
    expect(
      decodeChunks([gbkName.subarray(0, 20), gbkName.subarray(20)], 'win32', { codePage: 936 }),
    ).toBe('Microsoft Excel \u5de5\u4f5c\u8868')
  })

  it('returns an empty string for no chunks', () => {
    expect(decodeChunks([])).toBe('')
  })
})

describe('decodeUtf8 / decodeChunksUtf8', () => {
  it('does not throw on a UTF-8 tail truncated mid-character', () => {
    const cut = Buffer.from('{"t":"\u00e9\u00e9\u00e9\u4e2d\u4e2d\u4e2d', 'utf8').subarray(0, -1)
    expect(decodeUtf8(cut)).toContain('\uFFFD')
    expect(decodeChunksUtf8([cut])).toContain('\uFFFD')
  })

  it('decodes a multi-byte character split across chunks', () => {
    const bytes = Buffer.from('\u5de5\u4f5c\u8868', 'utf8')
    expect(decodeChunksUtf8([bytes.subarray(0, 2), bytes.subarray(2)])).toBe('\u5de5\u4f5c\u8868')
  })
})
