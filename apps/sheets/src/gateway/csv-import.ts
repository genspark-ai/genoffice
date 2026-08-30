/// CSV → minimal xlsx conversion for the open-file path. Values only: cells
/// that look like plain numbers become numeric, everything else stays text
/// (leading zeros survive). The converted file is a fresh workbook.

import JSZip from 'jszip'

const DELIMITERS = [',', ';', '\t'] as const

// Excel writes CSV in the system's legacy charset, not UTF-8 (GBK on Chinese
// Windows, Shift_JIS on Japanese), so decoding everything as UTF-8 turns every
// non-ASCII cell into replacement characters.
const LEGACY_CHARSETS = ['gb18030', 'shift_jis', 'big5', 'euc-kr', 'windows-1252'] as const

function decode(bytes: Uint8Array, charset: string, fatal = false): string | null {
  try {
    return new TextDecoder(charset, { fatal }).decode(bytes)
  } catch {
    return null
  }
}

/** how plausible a decoding is: CJK/ASCII good, replacement chars and control bytes bad */
function score(text: string): number {
  let value = 0
  for (const character of text) {
    const code = character.codePointAt(0) ?? 0
    if (character === '�') value -= 20
    else if (code < 0x20 && character !== '\n' && character !== '\r' && character !== '\t')
      value -= 10
    else if (code < 0x7f) value += 1
    // half-width katakana is the signature of double-byte text misread as
    // Shift_JIS, and near-absent from real spreadsheets — never a good sign
    else if (code >= 0xff61 && code <= 0xff9f) value -= 2
    else if (
      (code >= 0x3000 && code <= 0x9fff) || // CJK punctuation, kana, unified ideographs
      (code >= 0xac00 && code <= 0xd7af) || // hangul
      (code >= 0xff00 && code <= 0xffef) // full-width forms
    )
      value += 2
  }
  return value
}

/**
 * Decodes CSV bytes: BOM, then strict UTF-8, then the legacy charsets Excel
 * writes. `preferred` (from the UI language) breaks the ties those charsets
 * produce — GBK and Shift_JIS both decode the same bytes to plausible-looking
 * but different CJK.
 */
export function decodeCsvBuffer(bytes: Uint8Array, preferred?: string): string {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return decode(bytes.subarray(3), 'utf-8') ?? ''
  }
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return decode(bytes.subarray(2), 'utf-16le') ?? ''
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return decode(bytes.subarray(2), 'utf-16be') ?? ''

  const utf8 = decode(bytes, 'utf-8', true)
  if (utf8 !== null) return utf8

  let best = decode(bytes, 'utf-8') ?? ''
  let bestScore = score(best)
  const candidates = preferred ? [preferred, ...LEGACY_CHARSETS] : LEGACY_CHARSETS
  for (const charset of candidates) {
    const candidate = decode(bytes, charset)
    if (candidate === null) continue
    const candidateScore = score(candidate)
    if (candidateScore > bestScore) {
      best = candidate
      bestScore = candidateScore
    }
  }
  return best
}

/// Counts delimiter occurrences outside quotes over the first lines and
/// picks the most frequent one; ties favor the comma.
export function sniffDelimiter(text: string): string {
  const sample = text
    .slice(0, 64 * 1024)
    .split(/\r?\n/)
    .slice(0, 20)
  const counts = new Map<string, number>(DELIMITERS.map((d) => [d, 0]))
  for (const line of sample) {
    let quoted = false
    for (const character of line) {
      if (character === '"') quoted = !quoted
      else if (!quoted && counts.has(character)) {
        counts.set(character, (counts.get(character) ?? 0) + 1)
      }
    }
  }
  let best: string = DELIMITERS[0]
  let bestCount = -1
  for (const delimiter of DELIMITERS) {
    const count = counts.get(delimiter) ?? 0
    if (count > bestCount) {
      best = delimiter
      bestCount = count
    }
  }
  return best
}

export function parseCsv(input: string, delimiter = sniffDelimiter(input)): string[][] {
  const text = input.startsWith('﻿') ? input.slice(1) : input
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let quoted = false
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index]
    if (quoted) {
      if (character === '"') {
        if (text[index + 1] === '"') {
          field += '"'
          index += 1
        } else {
          quoted = false
        }
      } else {
        field += character
      }
      continue
    }
    if (character === '"' && field === '') {
      quoted = true
    } else if (character === delimiter) {
      row.push(field)
      field = ''
    } else if (character === '\n' || character === '\r') {
      if (character === '\r' && text[index + 1] === '\n') index += 1
      row.push(field)
      rows.push(row)
      row = []
      field = ''
    } else {
      field += character
    }
  }
  if (field !== '' || row.length > 0) {
    row.push(field)
    rows.push(row)
  }
  // A trailing newline produces one empty row — drop it.
  while (rows.length > 0 && rows[rows.length - 1]?.every((cell) => cell === '')) rows.pop()
  return rows
}

/// Plain decimal numbers only; leading zeros ("007") stay text so codes and
/// phone numbers survive the import.
export function isNumericCell(value: string): boolean {
  if (!/^-?(0|[1-9][0-9]*)(\.[0-9]+)?([eE][+-]?[0-9]+)?$/.test(value)) return false
  return Number.isFinite(Number(value))
}

function escapeXml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
}

function columnLabel(column: number): string {
  let label = ''
  let remaining = column + 1
  while (remaining > 0) {
    remaining -= 1
    label = String.fromCharCode(65 + (remaining % 26)) + label
    remaining = Math.floor(remaining / 26)
  }
  return label
}

export function buildWorksheetXml(rows: readonly (readonly string[])[]): string {
  const lines: string[] = []
  let maxColumns = 1
  rows.forEach((row, rowIndex) => {
    const cells: string[] = []
    row.forEach((value, columnIndex) => {
      if (value === '') return
      maxColumns = Math.max(maxColumns, columnIndex + 1)
      const reference = `${columnLabel(columnIndex)}${rowIndex + 1}`
      cells.push(
        isNumericCell(value)
          ? `<c r="${reference}"><v>${value}</v></c>`
          : `<c r="${reference}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`,
      )
    })
    if (cells.length > 0) lines.push(`<row r="${rowIndex + 1}">${cells.join('')}</row>`)
  })
  const dimension = `A1:${columnLabel(maxColumns - 1)}${Math.max(rows.length, 1)}`
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    `<dimension ref="${dimension}"/><sheetData>${lines.join('')}</sheetData></worksheet>`
  )
}

export async function csvToXlsxBuffer(csvText: string, sheetName = 'Sheet1'): Promise<Buffer> {
  const rows = parseCsv(csvText)
  if (rows.length === 0) throw new Error('The CSV file has no data rows.')
  return xlsxBufferFromRows(rows, sheetName)
}

/**
 * xlsx from the app's OWN comma-serialized sheet grid (AI create_document):
 * unlike the import path above, the delimiter is fixed to comma — cell text
 * may legitimately hold more semicolons/tabs than commas (csvField quotes
 * neither), and sniffing would then split the wrong columns — and an
 * all-empty grid becomes a valid blank workbook instead of an import error.
 */
export async function sheetCsvToXlsxBuffer(csvText: string, sheetName = 'Sheet1'): Promise<Buffer> {
  return xlsxBufferFromRows(parseCsv(csvText, ','), sheetName)
}

/** minimal empty workbook: the backing file for a "new blank spreadsheet" tab */
export async function blankXlsxBuffer(sheetName = 'Sheet1'): Promise<Buffer> {
  return xlsxBufferFromRows([], sheetName)
}

async function xlsxBufferFromRows(
  rows: readonly (readonly string[])[],
  sheetName: string,
): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `<sheets><sheet name="${escapeXml(sheetName)}" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '</Relationships>',
  )
  zip.file('xl/worksheets/sheet1.xml', buildWorksheetXml(rows))
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}
