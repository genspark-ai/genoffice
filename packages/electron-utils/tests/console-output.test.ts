import { describe, expect, it } from 'vitest'
import { decodeChunks, decodeCommandOutput } from '../src/console-output'

// GBK bytes for a Chinese spreadsheet term, as emitted by reg.exe and
// powershell.exe on a Chinese Windows install.
const gbkName = Buffer.concat([
  Buffer.from('Microsoft Excel ', 'ascii'),
  Buffer.from([0xb9, 0xa4, 0xd7, 0xf7, 0xb1, 0xed]),
])

describe('decodeCommandOutput', () => {
  it('recovers OEM code page text on Windows instead of replacement boxes', () => {
    expect(gbkName.toString('utf8')).toContain('\uFFFD')
    expect(decodeCommandOutput(gbkName, 'win32')).toBe('Microsoft Excel 工作表')
  })

  it('leaves ASCII output untouched on Windows', () => {
    const ascii = Buffer.from('    (Default)    REG_SZ    Word.Document.12\r\n', 'ascii')
    expect(decodeCommandOutput(ascii, 'win32')).toBe(ascii.toString('utf8'))
  })

  it('keeps genuine UTF-8 output on Windows', () => {
    const utf8 = Buffer.from('Microsoft PowerPoint 演示文稿', 'utf8')
    expect(decodeCommandOutput(utf8, 'win32')).toBe('Microsoft PowerPoint 演示文稿')
  })

  it('honours UTF-16 BOMs emitted by redirected PowerShell', () => {
    const text = '证书 Subject'
    const le = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(text, 'utf16le')])
    expect(decodeCommandOutput(le, 'win32')).toBe(text)
    const be = Buffer.concat([Buffer.from([0xfe, 0xff]), Buffer.from(text, 'utf16le').swap16()])
    expect(decodeCommandOutput(be, 'win32')).toBe(text)
  })

  it('uses UTF-8 on macOS and Linux', () => {
    const utf8 = Buffer.from('WPS Office — 表格', 'utf8')
    expect(decodeCommandOutput(utf8, 'linux')).toBe('WPS Office — 表格')
    expect(decodeCommandOutput(utf8, 'darwin')).toBe('WPS Office — 表格')
  })

  it('passes through an already-decoded string', () => {
    expect(decodeCommandOutput('Word Document', 'win32')).toBe('Word Document')
  })
})

describe('decodeChunks', () => {
  it('decodes a multi-byte character split across two chunks', () => {
    const bytes = Buffer.from('工作表', 'utf8')
    const cut = 2
    const chunks = [bytes.subarray(0, cut), bytes.subarray(cut)]
    expect(chunks.map((c) => c.toString()).join('')).toContain('\uFFFD')
    expect(decodeChunks(chunks)).toBe('工作表')
  })

  it('joins string chunks without re-encoding them', () => {
    expect(decodeChunks(['Microsoft ', 'Excel 工作表'])).toBe('Microsoft Excel 工作表')
  })

  it('decodes a Windows OEM stream as one unit', () => {
    expect(decodeChunks([gbkName.subarray(0, 20), gbkName.subarray(20)], 'win32')).toBe(
      'Microsoft Excel 工作表',
    )
  })

  it('returns an empty string for no chunks', () => {
    expect(decodeChunks([])).toBe('')
  })
})
