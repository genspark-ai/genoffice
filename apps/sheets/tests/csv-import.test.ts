import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  blankXlsxBuffer,
  buildWorksheetXml,
  csvToXlsxBuffer,
  sheetCsvToXlsxBuffer,
  decodeCsvBuffer,
  isNumericCell,
  parseCsv,
  sniffDelimiter,
} from '../src/gateway/csv-import'

describe('decodeCsvBuffer', () => {
  const rows = '城市,人口\n东京,37\n'
  const jp = '都市,人口\n東京,37\n'

  it('reads UTF-8 with and without a BOM', () => {
    expect(decodeCsvBuffer(Buffer.from(rows, 'utf8'))).toBe(rows)
    expect(
      decodeCsvBuffer(Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(rows, 'utf8')])),
    ).toBe(rows)
  })

  it('reads UTF-16 in both byte orders', () => {
    expect(decodeCsvBuffer(Buffer.from(`\uFEFF${rows}`, 'utf16le'))).toBe(rows)
    const be = Buffer.from(`\uFEFF${rows}`, 'utf16le').swap16()
    expect(decodeCsvBuffer(be)).toBe(rows)
  })

  it('falls back to the legacy charset Excel actually writes', () => {
    expect(decodeCsvBuffer(gbkBytes(rows))).toBe(rows)
    expect(decodeCsvBuffer(shiftJisBytes(jp), 'shift_jis')).toBe(jp)
  })

  it('uses the preferred charset to break GBK/Shift_JIS ties', () => {
    // these bytes decode to plausible CJK under both, so the UI language decides
    expect(decodeCsvBuffer(shiftJisBytes(jp))).not.toBe(jp)
    expect(decodeCsvBuffer(gbkBytes(rows), 'gb18030')).toBe(rows)
  })

  it('keeps plain ASCII untouched', () => {
    expect(decodeCsvBuffer(Buffer.from('a,b\n1,2\n', 'utf8'))).toBe('a,b\n1,2\n')
  })
})

/** encodes via the platform decoder tables by round-tripping every code point */
function encodeWith(text: string, charset: string): Buffer {
  // Node has no legacy encoder, so build the byte string from a known mapping by
  // brute-forcing each character against the decoder.
  const out: number[] = []
  const decoder = new TextDecoder(charset)
  for (const character of text) {
    if (character.codePointAt(0)! < 0x80) {
      out.push(character.codePointAt(0)!)
      continue
    }
    let found = false
    for (let lead = 0x81; lead <= 0xfe && !found; lead += 1) {
      for (let trail = 0x40; trail <= 0xfe && !found; trail += 1) {
        if (decoder.decode(new Uint8Array([lead, trail])) === character) {
          out.push(lead, trail)
          found = true
        }
      }
    }
    if (!found) throw new Error(`cannot encode ${character} in ${charset}`)
  }
  return Buffer.from(out)
}

const gbkBytes = (text: string): Buffer => encodeWith(text, 'gb18030')
const shiftJisBytes = (text: string): Buffer => encodeWith(text, 'shift_jis')

describe('parseCsv', () => {
  it('handles quotes, embedded delimiters, escaped quotes, and CRLF', () => {
    const rows = parseCsv('a,"b,1","say ""hi""",c\r\nd,,e,')
    expect(rows).toEqual([
      ['a', 'b,1', 'say "hi"', 'c'],
      ['d', '', 'e', ''],
    ])
  })

  it('keeps newlines inside quoted fields and strips the BOM', () => {
    const rows = parseCsv('﻿name,note\nx,"line1\nline2"')
    expect(rows).toEqual([
      ['name', 'note'],
      ['x', 'line1\nline2'],
    ])
  })

  it('sniffs semicolon and tab delimiters', () => {
    expect(sniffDelimiter('a;b;c\n1;2;3')).toBe(';')
    expect(sniffDelimiter('a\tb\tc\n1\t2\t3')).toBe('\t')
    expect(parseCsv('a;b\n1;2')).toEqual([
      ['a', 'b'],
      ['1', '2'],
    ])
  })

  it('drops the trailing empty row from a final newline', () => {
    expect(parseCsv('a,b\n1,2\n')).toEqual([
      ['a', 'b'],
      ['1', '2'],
    ])
  })
})

describe('isNumericCell', () => {
  it('accepts plain decimals and scientific notation', () => {
    for (const value of ['0', '42', '-3.5', '1e3', '0.5']) {
      expect(isNumericCell(value), value).toBe(true)
    }
  })

  it('keeps codes, dates, and padded numbers as text', () => {
    for (const value of ['007', '2025-06-01', '1,234', '+86', '', ' 5']) {
      expect(isNumericCell(value), value).toBe(false)
    }
  })
})

describe('csvToXlsxBuffer', () => {
  it('produces a workbook whose sheet carries typed cells', async () => {
    const buffer = await csvToXlsxBuffer('Item,Qty\n"A, large",007\nB,42')
    const zip = await JSZip.loadAsync(buffer)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(sheet).toContain(
      '<c r="A2" t="inlineStr"><is><t xml:space="preserve">A, large</t></is></c>',
    )
    expect(sheet).toContain('<c r="B2" t="inlineStr"><is><t xml:space="preserve">007</t></is></c>')
    expect(sheet).toContain('<c r="B3"><v>42</v></c>')
    expect(await zip.file('xl/workbook.xml')?.async('text')).toContain('<sheet name="Sheet1"')
  })

  it('rejects an empty file and escapes XML metacharacters', async () => {
    await expect(csvToXlsxBuffer('')).rejects.toThrow('no data rows')
    expect(buildWorksheetXml([['<b>&"']])).toContain('&lt;b&gt;&amp;&quot;')
  })
})

describe('blankXlsxBuffer', () => {
  it('produces a one-sheet workbook with an empty grid', async () => {
    const zip = await JSZip.loadAsync(await blankXlsxBuffer())
    expect(await zip.file('xl/workbook.xml')?.async('text')).toContain('<sheet name="Sheet1"')
    const sheet = await zip.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(sheet).toContain('<dimension ref="A1:A1"/><sheetData></sheetData>')
  })
})

describe('sheetCsvToXlsxBuffer (AI create_document)', () => {
  it('never sniffs the delimiter: semicolon-heavy cells stay one column', async () => {
    // csvField quotes neither ; nor \t — sniffing would split these cells
    const buffer = await sheetCsvToXlsxBuffer('a;b;c,x\r\nd;e;f,y\r\n', 'de')
    const zip = await JSZip.loadAsync(buffer)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(sheet).toContain('<t xml:space="preserve">a;b;c</t>')
    expect(sheet).toContain('<c r="B1"')
    expect(sheet).not.toContain('<c r="C1"')
    expect(await zip.file('xl/workbook.xml')?.async('text')).toContain('<sheet name="de"')
  })

  it('turns an all-empty grid into a valid blank workbook', async () => {
    const zip = await JSZip.loadAsync(await sheetCsvToXlsxBuffer('\r\n', 'Empty'))
    const sheet = await zip.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(sheet).toContain('<sheetData></sheetData>')
    expect(await zip.file('xl/workbook.xml')?.async('text')).toContain('<sheet name="Empty"')
  })
})
