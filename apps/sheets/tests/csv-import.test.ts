import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  blankXlsxBuffer,
  buildWorksheetXml,
  csvToXlsxBuffer,
  csvToXlsxBufferForOpen,
  sheetCsvToXlsxBuffer,
  decodeCsvBuffer,
  isNumericCell,
  parseCsv,
  resolveImportDelimiter,
  sniffDelimiter,
  splitSepDeclaration,
} from '@genoffice/xlsx-gateway/gateway/csv-import'

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

  it('reads BOM-less UTF-16 instead of keeping NUL garbage', () => {
    const ascii = 'a,b\n1,2\n'
    const le = Buffer.from(ascii, 'utf16le')
    expect(decodeCsvBuffer(le)).toBe(ascii)
    const be = Buffer.from(ascii, 'utf16le').swap16()
    expect(decodeCsvBuffer(be)).toBe(ascii)
    const leRows = Buffer.from(rows, 'utf16le')
    expect(decodeCsvBuffer(leRows)).toBe(rows)
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

describe('sep= declaration', () => {
  it('names the delimiter and is never a data row', () => {
    const text = 'sep=;\nname;qty\nApple;3\n'
    expect(splitSepDeclaration(text)).toEqual({ text: 'name;qty\nApple;3\n', delimiter: ';' })
    expect(sniffDelimiter(text)).toBe(';')
    expect(parseCsv(text)).toEqual([
      ['name', 'qty'],
      ['Apple', '3'],
    ])
    expect(sniffDelimiter('\ufeffsep=\t\na\tb\n')).toBe('\t')
    expect(splitSepDeclaration('separator,x\n1,2\n').delimiter).toBeUndefined()
  })

  it('overrides the prose guard on the import path', () => {
    const prose = 'sep=;\nnotes\nhello; world\nplain text\n'
    expect(resolveImportDelimiter(prose)).toBe(';')
    expect(resolveImportDelimiter(prose.slice('sep=;\n'.length))).toBe(',')
  })

  it('is stripped by the workbook import too', async () => {
    const zip = await JSZip.loadAsync(await csvToXlsxBuffer('sep=;\na;b\n1;2\n'))
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).not.toContain('sep=')
    expect(sheet).toContain('<c r="B2"><v>2</v></c>')
  })
})

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

  it('ignores delimiters inside escaped quotes when sniffing', () => {
    // The "" escape keeps the field quoted: all five inner ; must not count,
    // or they outvote the two true commas and the columns mis-split.
    expect(sniffDelimiter('a,"; ""; ""; ""; """,d')).toBe(',')
    expect(parseCsv('a,"; ""; ""; ""; """,d')).toEqual([['a', '; "; "; "; "', 'd']])
  })

  it('ignores delimiters inside multiline quoted fields when sniffing', () => {
    // The quoted field spans lines: without carried quote state the three
    // inner ; would outvote the two true commas and the columns mis-split.
    const text = 'a,b\n"x\ny;z;w;v",q\n'
    expect(sniffDelimiter(text)).toBe(',')
    expect(parseCsv(text, ',')).toEqual([
      ['a', 'b'],
      ['x\ny;z;w;v', 'q'],
    ])
  })

  it('drops the trailing empty row from a final newline', () => {
    expect(parseCsv('a,b\n1,2\n')).toEqual([
      ['a', 'b'],
      ['1', '2'],
    ])
  })
})

describe('resolveImportDelimiter', () => {
  it('keeps single-column prose with stray semicolons in one column', () => {
    const notes = 'Notes\nhello; world\nfoo; bar; baz\n'
    expect(sniffDelimiter(notes)).toBe(';')
    expect(resolveImportDelimiter(notes)).toBe(',')
    expect(parseCsv(notes, resolveImportDelimiter(notes))).toEqual([
      ['Notes'],
      ['hello; world'],
      ['foo; bar; baz'],
    ])
  })

  it('keeps genuine semicolon tables split', () => {
    expect(resolveImportDelimiter('a;b;c\n1;2;3')).toBe(';')
    expect(resolveImportDelimiter('a;b\nc')).toBe(';')
  })

  it('keeps a comma-free table split when a title row precedes uniform body rows', () => {
    expect(resolveImportDelimiter('Sales 2026\na;b;c\n1;2;3\n4;5;6')).toBe(';')
    expect(resolveImportDelimiter('Report\n\na\tb\n1\t2\n3\t4\n')).toBe('\t')
    // one wide row among prose lines is not a table
    expect(resolveImportDelimiter('Notes\nhello; world\nfoo bar\nbaz qux\nx; y')).toBe(',')
  })

  it('keeps the sniffed delimiter when comma is equally ragged', () => {
    // header opens single-field under ';' but commas appear irregularly too:
    // forcing comma would trade one mis-split for another, so stay put
    expect(resolveImportDelimiter('Notes\nhello; world; x, y\nfoo;bar')).toBe(';')
  })
})

describe('isNumericCell', () => {
  it('accepts plain decimals and scientific notation', () => {
    for (const value of ['0', '42', '-3.5', '1e3', '0.5', '.5', '1.', '-.5']) {
      expect(isNumericCell(value), value).toBe(true)
    }
  })

  it('keeps codes, dates, and padded numbers as text', () => {
    for (const value of ['007', '2025-06-01', '1,234', '+86', '', ' 5']) {
      expect(isNumericCell(value), value).toBe(false)
    }
  })

  it('keeps integers past Excel precision as text so long IDs survive', () => {
    expect(isNumericCell('123456789012345')).toBe(true)
    expect(isNumericCell('1234567890123456')).toBe(false)
    expect(isNumericCell('12345678901234567890')).toBe(false)
    expect(buildWorksheetXml([['12345678901234567890']])).toContain('t="inlineStr"')
    for (const value of [
      '0.3333333333333333',
      '1.4142135623730951',
      '0.30000000000000004',
      '3.14159265358979',
    ]) {
      expect(isNumericCell(value), value).toBe(true)
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
    expect(buildWorksheetXml([['a\rb_x000D_']])).toContain('>a_x000D_b_x005F_x000D_<')
  })

  it('rejects worksheet names that Excel forbids', async () => {
    for (const name of ['data[final]', 'bad/name', "'quoted", "trailing'"]) {
      await expect(csvToXlsxBuffer('A\n1', name), name).rejects.toThrow(/sheet name/i)
      await expect(blankXlsxBuffer(name), name).rejects.toThrow(/sheet name/i)
    }
  })
})

describe('csvToXlsxBufferForOpen', () => {
  it('opens a blank-lines-only CSV as an empty workbook and reports the empty source', async () => {
    const { buffer, empty } = await csvToXlsxBufferForOpen('\r\n'.repeat(5_927))
    const zip = await JSZip.loadAsync(buffer)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')?.async('text')

    expect(empty).toBe(true)
    expect(sheet).toContain('<sheetData></sheetData>')
  })

  it('keeps normal CSV rows unchanged', async () => {
    const { buffer, empty } = await csvToXlsxBufferForOpen('name,amount\r\nalpha,10\r\n')
    const zip = await JSZip.loadAsync(buffer)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')?.async('text')

    expect(empty).toBe(false)
    expect(sheet).toContain('<t xml:space="preserve">alpha</t>')
    expect(sheet).toContain('<v>10</v>')
  })

  it('honors a pinned delimiter so a comma-heavy .tsv keeps its columns', async () => {
    // A UniProt-style export (issue #1140): few columns, and each annotation
    // field is itself a comma-separated list, so the sniffer counts more
    // commas than tabs and shatters every row. The .tsv extension already
    // declares the delimiter — pin it.
    const tsv =
      'Entry\tProtein names\tGene names\r\n' +
      'P38398\tBreast cancer protein 1, RING finger E3 ubiquitin protein ligase, BRC1, BRCA1, PARSUM1, PPP2R5S\tBRCA1, RNF3, c.1125dupT, 1253del7TTGGTCTAA\r\n' +
      'Q92766\tSentrin/SUMO-specific protease 6, GCP-2, SUMO-1/sentrin-specific peptidase activity, PIASY, SENP6\tSENP6, SENP7, GCP2, CA1H5\r\n'
    expect(sniffDelimiter(tsv)).toBe(',')
    const lastColumn = async (buffer: Buffer): Promise<string> => {
      const zip = await JSZip.loadAsync(buffer)
      const sheet = await zip.file('xl/worksheets/sheet1.xml')?.async('text')
      return /<dimension ref="A1:([A-Z]+)/.exec(sheet ?? '')?.[1] ?? ''
    }

    const pinned = await csvToXlsxBufferForOpen(tsv, 'Sheet1', '\t')
    const sniffed = await csvToXlsxBufferForOpen(tsv)

    // three tab columns instead of nine comma fragments, each field whole
    expect(pinned.empty).toBe(false)
    expect(await lastColumn(pinned.buffer)).toBe('C')
    expect(await lastColumn(sniffed.buffer)).toBe('I')
    const sheet = await (
      await JSZip.loadAsync(pinned.buffer)
    )
      .file('xl/worksheets/sheet1.xml')
      ?.async('text')
    expect(sheet).toContain(
      '<t xml:space="preserve">BRCA1, RNF3, c.1125dupT, 1253del7TTGGTCTAA</t>',
    )
    expect(sheet).toContain('<t xml:space="preserve">P38398</t>')
    // a blank .tsv still opens as an empty workbook instead of throwing
    expect((await csvToXlsxBufferForOpen('\r\n\r\n', 'Sheet1', '\t')).empty).toBe(true)
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
