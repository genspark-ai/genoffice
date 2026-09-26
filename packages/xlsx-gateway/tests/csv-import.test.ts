import { describe, expect, it } from 'vitest'
import {
  decodeCsvBuffer,
  isNumericCell,
  MAX_CSV_COLS,
  MAX_CSV_ROWS,
  parseCsv,
  resolveImportDelimiter,
  sniffDelimiter,
  splitSepDeclaration,
} from '../src/gateway/csv-import'

function utf16leBytes(text: string): Uint8Array {
  const bytes = new Uint8Array(text.length * 2)
  for (let i = 0; i < text.length; i += 1) {
    const code = text.charCodeAt(i)
    bytes[i * 2] = code & 0xff
    bytes[i * 2 + 1] = (code >> 8) & 0xff
  }
  return bytes
}

describe('parseCsv with explicit delimiters', () => {
  it('parses a simple comma grid', () => {
    expect(parseCsv('a,b,c\n1,2,3', ',')).toEqual([
      ['a', 'b', 'c'],
      ['1', '2', '3'],
    ])
  })

  it('parses semicolon and tab grids', () => {
    expect(parseCsv('a;b\n1;2', ';')).toEqual([
      ['a', 'b'],
      ['1', '2'],
    ])
    expect(parseCsv('a\tb\n1\t2', '\t')).toEqual([
      ['a', 'b'],
      ['1', '2'],
    ])
  })

  it('keeps delimiters inside quoted fields', () => {
    expect(parseCsv('"a,b",c\n1,2,3', ',')).toEqual([
      ['a,b', 'c'],
      ['1', '2', '3'],
    ])
  })

  it('unescapes doubled quotes', () => {
    expect(parseCsv('"a""b",c', ',')).toEqual([['a"b', 'c']])
  })

  it('keeps line breaks inside quoted fields', () => {
    expect(parseCsv('"a\nb",c\n1,2', ',')).toEqual([
      ['a\nb', 'c'],
      ['1', '2'],
    ])
  })

  it('handles CRLF and drops the trailing empty row', () => {
    expect(parseCsv('a,b\r\n1,2\r\n', ',')).toEqual([
      ['a', 'b'],
      ['1', '2'],
    ])
  })

  it('returns an empty grid for empty input', () => {
    expect(parseCsv('', ',')).toEqual([])
  })

  it('strips a sep declaration line before parsing', () => {
    expect(parseCsv('sep=;\na;b\nc;d')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ])
  })
})

describe('sniffDelimiter', () => {
  it('picks the most frequent delimiter and favors comma on ties', () => {
    expect(sniffDelimiter('a,b,c\n1,2,3')).toBe(',')
    expect(sniffDelimiter('a;b;c\n1;2;3')).toBe(';')
    expect(sniffDelimiter('a\tb\tc\n1\t2\t3')).toBe('\t')
    expect(sniffDelimiter('plain text without delimiters')).toBe(',')
  })

  it('ignores delimiters inside quoted fields', () => {
    expect(sniffDelimiter('"a;b",c\n"d;e",f')).toBe(',')
  })

  it('honors a sep declaration line', () => {
    expect(sniffDelimiter('sep=;\na,b;c')).toBe(';')
    expect(sniffDelimiter('SEP=,\na;b,c')).toBe(',')
  })
})

describe('splitSepDeclaration', () => {
  it('splits the declaration from the body', () => {
    expect(splitSepDeclaration('sep=;\na;b')).toEqual({ text: 'a;b', delimiter: ';' })
  })

  it('returns the input unchanged without a declaration', () => {
    expect(splitSepDeclaration('a,b\n1,2')).toEqual({ text: 'a,b\n1,2' })
  })
})

describe('resolveImportDelimiter', () => {
  it('uses an explicit sep declaration', () => {
    expect(resolveImportDelimiter('sep=;\na;b\nc;d')).toBe(';')
  })

  it('keeps comma for comma tables', () => {
    expect(resolveImportDelimiter('a,b\n1,2')).toBe(',')
  })

  it('keeps semicolon for a multi-field table', () => {
    expect(resolveImportDelimiter('a;b\n1;2\n3;4')).toBe(';')
  })

  it('treats prose with stray separators as a single comma column', () => {
    expect(resolveImportDelimiter('Title\nhello; world\njust note')).toBe(',')
  })
})

describe('decodeCsvBuffer', () => {
  it('passes ASCII bytes through', () => {
    const bytes = new TextEncoder().encode('a,b\n1,2')
    expect(decodeCsvBuffer(bytes)).toBe('a,b\n1,2')
  })

  it('strips a UTF-8 BOM', () => {
    const body = new TextEncoder().encode('a,b')
    const bytes = new Uint8Array([0xef, 0xbb, 0xbf, ...body])
    expect(decodeCsvBuffer(bytes)).toBe('a,b')
  })

  it('decodes UTF-16LE with a BOM', () => {
    const bytes = new Uint8Array([0xff, 0xfe, ...utf16leBytes('a,b\n1,2')])
    expect(decodeCsvBuffer(bytes)).toBe('a,b\n1,2')
  })

  it('decodes UTF-16LE without a BOM from the NUL pattern', () => {
    expect(decodeCsvBuffer(utf16leBytes('a,b'))).toBe('a,b')
  })
})

describe('parseCsv row/col caps', () => {
  it('rejects too many columns, counting the last field of an unterminated row', () => {
    const wide = Array(MAX_CSV_COLS + 1)
      .fill('a')
      .join(',')
    expect(() => parseCsv(wide, ',')).toThrow(/too many columns/)
    expect(parseCsv(Array(MAX_CSV_COLS).fill('a').join(','), ',')[0]).toHaveLength(MAX_CSV_COLS)
  })

  it('accepts normal grids', () => {
    expect(parseCsv('a,b\n1,2', ',')).toHaveLength(2)
  })

  it('exposes row/col caps', () => {
    expect(MAX_CSV_ROWS).toBe(1_048_576)
    expect(MAX_CSV_COLS).toBe(16_384)
  })
})

describe('isNumericCell', () => {
  it('accepts plain decimal numbers', () => {
    expect(isNumericCell('123')).toBe(true)
    expect(isNumericCell('-12.5')).toBe(true)
    expect(isNumericCell('1e10')).toBe(true)
    expect(isNumericCell('0')).toBe(true)
  })

  it('accepts Excel-style bare-dot decimals but keeps a leading plus as text', () => {
    for (const value of ['.5', '1.', '-.5', '-1.', '.5e2']) {
      expect(isNumericCell(value), value).toBe(true)
    }
    for (const value of ['+1', '+86', '.', '-.', '-', '1..5', '.5.']) {
      expect(isNumericCell(value), value).toBe(false)
    }
  })

  it('rejects leading-zero codes so they stay text', () => {
    expect(isNumericCell('007')).toBe(false)
    expect(isNumericCell('0123')).toBe(false)
  })

  it('rejects integers past Excel precision so long ids stay text', () => {
    expect(isNumericCell('123456789012345')).toBe(true)
    expect(isNumericCell('1234567890123456')).toBe(false)
  })

  it('rejects blanks and formatted text', () => {
    expect(isNumericCell('')).toBe(false)
    expect(isNumericCell('12a')).toBe(false)
    expect(isNumericCell('1,000')).toBe(false)
    expect(isNumericCell(' 123')).toBe(false)
  })
})
