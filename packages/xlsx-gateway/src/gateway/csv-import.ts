/// CSV parsing for the import path. Values only: cells that look like plain
/// numbers become numeric, everything else stays text (leading zeros survive).
///
/// The sheets renderer imports this module (Data › From Text/CSV), so it stays
/// free of node APIs: the half that writes the xlsx archive lives in
/// csv-import-xlsx, which only the desktop path loads.

import { encodeXlsxEscapes } from './xlsx-escapes'

const DELIMITERS = [',', ';', '\t'] as const

/** Excel's sheet bounds; anything past them cannot land in a workbook anyway. */
export const MAX_CSV_ROWS = 1_048_576
export const MAX_CSV_COLS = 16_384

/// Default byte ceiling for an in-cell CSV import (genoffice#2505): 128 MiB.
/// Mirrors DEFAULT_CSV_IMPORT_MAX_BYTES in @genoffice/electron-utils (the
/// main-process open path's configurable cap) — this renderer-side copy is
/// the default only; the workbook-open path reads the configured value.
export const DEFAULT_CSV_IMPORT_MAX_BYTES = 128 * 1024 * 1024

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
/**
 * UTF-16 without a BOM (Excel writes it; editors strip it): NUL bytes then
 * interleave the text, which strict UTF-8 accepts — so this must run before
 * the UTF-8 attempt or the NULs survive as garbage. NUL bytes never occur in
 * UTF-8/legacy CSV text (GBK/Shift_JIS trail bytes exclude 0x00), so any
 * meaningful NUL presence means UTF-16, with the NUL-heavy side picking the
 * byte order (ties go LE, Excel's order).
 */
function sniffUtf16WithoutBom(bytes: Uint8Array): 'utf-16le' | 'utf-16be' | null {
  const sample = bytes.subarray(0, Math.min(bytes.length, 1024))
  let pairs = 0
  let evenNul = 0
  let oddNul = 0
  for (let i = 0; i + 1 < sample.length; i += 2) {
    pairs += 1
    if (sample[i] === 0) evenNul += 1
    if (sample[i + 1] === 0) oddNul += 1
  }
  if (pairs < 2) return null
  if ((evenNul + oddNul) / (pairs * 2) < 0.05) return null
  return oddNul >= evenNul ? 'utf-16le' : 'utf-16be'
}

export function decodeCsvBuffer(bytes: Uint8Array, preferred?: string): string {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return decode(bytes.subarray(3), 'utf-8') ?? ''
  }
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return decode(bytes.subarray(2), 'utf-16le') ?? ''
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return decode(bytes.subarray(2), 'utf-16be') ?? ''
  const bomless = sniffUtf16WithoutBom(bytes)
  if (bomless) return decode(bytes, bomless) ?? ''

  const utf8 = decode(bytes, 'utf-8', true)
  if (utf8 !== null) return utf8

  let best = decode(bytes, 'utf-8') ?? ''
  let bestScore = score(best)
  const utf16 = ['utf-16le', 'utf-16be'] as const
  const candidates = preferred ? [preferred, ...LEGACY_CHARSETS] : LEGACY_CHARSETS
  // CJK-only UTF-16 carries too few NUL bytes for the sniffer above; let the scorer pick.
  for (const charset of [...utf16, ...candidates]) {
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

/// Excel's own hint line: a first line of exactly `sep=<char>` names the
/// delimiter and is not data. Excel writes it for CSV exports in locales
/// that use ";" and honours it on open.
const SEP_DECLARATION = /^\uFEFF?sep=(.)\r?\n/i

export function splitSepDeclaration(text: string): { text: string; delimiter?: string } {
  const match = SEP_DECLARATION.exec(text)
  if (!match) return { text }
  return { text: text.slice(match[0].length), delimiter: match[1]! }
}

/// Counts delimiter occurrences outside quotes over the first lines and
/// picks the most frequent one; ties favor the comma. A `sep=` line wins.
export function sniffDelimiter(input: string): string {
  const { text, delimiter } = splitSepDeclaration(input)
  if (delimiter !== undefined) return delimiter
  const sample = text
    .slice(0, 64 * 1024)
    .split(/\r?\n/)
    .slice(0, 20)
  const counts = new Map<string, number>(DELIMITERS.map((d) => [d, 0]))
  // Quote state carries across lines: a quoted field may span line breaks,
  // and delimiters inside it must never be counted.
  let quoted = false
  for (const line of sample) {
    for (let index = 0; index < line.length; index += 1) {
      const character = line[index]
      if (character === undefined) continue
      if (character === '"') {
        // An escaped quote ("") stays inside the quoted field; only a lone
        // quote toggles quoting, mirroring parseCsv below.
        if (quoted && line[index + 1] === '"') {
          index += 1
        } else {
          quoted = !quoted
        }
      } else if (!quoted && counts.has(character)) {
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

/**
 * Where an import is: 'read' is the text being parsed, 'write' is the
 * workbook being written. The two are separate because the write is where
 * the time goes.
 */
export type CsvImportPhase = 'read' | 'write'

export interface CsvImportProgress {
  readonly phase: CsvImportPhase
  /** 0..1 within the phase */
  readonly fraction: number
}

export interface CsvImportOptions {
  /** called as the import advances */
  readonly onProgress?: (progress: CsvImportProgress) => void
  /** aborts the import; the conversion throws */
  readonly signal?: AbortSignal
}

/** how often the scan reports its position, in characters */
const PROGRESS_INTERVAL = 256 * 1024

/**
 * The grid's rows, one at a time.
 *
 * The import never holds the whole grid: a 40 MB CSV is millions of cells, and
 * a `string[][]` of them costs more than the file itself — the worksheet XML
 * they become costs several times more again. Callers that only need the
 * extent count the rows and drop them; the conversion streams them.
 *
 * `onOffset` receives the character offset reached (for progress) and
 * `signal` aborts the scan.
 */
export function* csvRows(
  input: string,
  delimiter: string,
  options?: { readonly onOffset?: (chars: number) => void; readonly signal?: AbortSignal },
): Generator<string[]> {
  const stripped = splitSepDeclaration(input).text
  const text = stripped.startsWith('﻿') ? stripped.slice(1) : stripped
  // An abort that arrived before the work started must not be ignored: the
  // scan would otherwise run to the end for a file small enough never to
  // reach a progress checkpoint.
  options?.signal?.throwIfAborted()
  const fail = (msg: string): never => {
    throw new Error(`CSV import rejected: ${msg}`)
  }
  let row: string[] = []
  let field = ''
  // A trailing newline produces one empty row — drop it. Empty rows are held
  // back until a row with content proves they were trailing.
  let held: string[][] = []
  let rows = 0
  const pushField = (): void => {
    row.push(field)
    if (row.length > MAX_CSV_COLS) fail(`too many columns (cap ${MAX_CSV_COLS})`)
    field = ''
  }
  let quoted = false
  let nextReport = PROGRESS_INTERVAL
  for (let index = 0; index < text.length; index += 1) {
    if (index >= nextReport) {
      options?.signal?.throwIfAborted()
      options?.onOffset?.(index)
      nextReport = index + PROGRESS_INTERVAL
    }
    const character = text[index]
    if (quoted) {
      if (character === '"') {
        if (text[index + 1] === '"') {
          field += '"'
          index += 1
        } else {
          quoted = false
        }
      } else if (character === '\r') {
        if (text[index + 1] === '\n') index += 1
        field += '\n'
      } else {
        field += character
      }
      continue
    }
    if (character === '"' && field === '') {
      quoted = true
    } else if (character === delimiter) {
      pushField()
    } else if (character === '\n' || character === '\r') {
      if (character === '\r' && text[index + 1] === '\n') index += 1
      pushField()
      rows += 1
      if (rows > MAX_CSV_ROWS) fail(`too many rows (cap ${MAX_CSV_ROWS})`)
      const finished = row
      row = []
      if (finished.every((cell) => cell === '')) {
        held.push(finished)
        continue
      }
      for (const pending of held) yield pending
      held = []
      yield finished
    } else {
      field += character
    }
  }
  if (field !== '' || row.length > 0) {
    pushField()
    rows += 1
    if (rows > MAX_CSV_ROWS) fail(`too many rows (cap ${MAX_CSV_ROWS})`)
    const finished = row
    // trailing empty rows are dropped: the held ones and this one
    if (!finished.every((cell) => cell === '')) {
      for (const pending of held) yield pending
      yield finished
    }
  }
}

export function parseCsv(input: string, delimiter = sniffDelimiter(input)): string[][] {
  return [...csvRows(input, delimiter)]
}

/// Plain decimal numbers only (".5", "1." and "-.5" count, as in Excel); leading
/// zeros ("007") and a "+" sign ("+86") stay text so codes and phone numbers
/// survive the import. Integers past Excel's 15-digit precision stay text too,
/// so long IDs are not corrupted on open.
export function isNumericCell(value: string): boolean {
  if (!/^-?(?:(?:0|[1-9][0-9]*)(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?$/.test(value))
    return false
  if (!/[.eE]/.test(value) && value.replace(/^-/, '').length > 15) return false
  return Number.isFinite(Number(value))
}

/** the XML escapes a cell text needs; shared with the archive writer */
export function escapeXml(value: string): string {
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

/** How far the grid reaches: how many rows, and how many columns carry a value. */
export interface CsvGridShape {
  readonly rows: number
  readonly columns: number
}

/** The widest column that holds a value, counted like the worksheet's
 *  dimension: an empty trailing cell does not make the sheet wider. */
function noteColumn(columns: number, row: readonly string[]): number {
  for (let index = row.length - 1; index >= 0; index -= 1) {
    if (row[index] !== '') return Math.max(columns, index + 1)
  }
  return columns
}

/** The shape of a grid that is already in memory. */
function shapeOfRows(rows: readonly (readonly string[])[]): CsvGridShape {
  let columns = 1
  for (const row of rows) columns = noteColumn(columns, row)
  return { rows: rows.length, columns }
}

/**
 * The grid's extent, read from the text in one pass: the rows are counted and
 * dropped. The worksheet's dimension has to name the extent before the first
 * row is written, so the conversion reads the shape first and streams the rows
 * after it.
 */
export function csvGridShape(
  input: string,
  delimiter: string,
  options?: CsvImportOptions,
): CsvGridShape {
  let rows = 0
  let columns = 1
  const total = Math.max(input.length, 1)
  const scan = {
    onOffset: (chars: number) => options?.onProgress?.({ phase: 'read', fraction: chars / total }),
    ...(options?.signal ? { signal: options.signal } : {}),
  }
  for (const row of csvRows(input, delimiter, scan)) {
    rows += 1
    columns = noteColumn(columns, row)
  }
  options?.onProgress?.({ phase: 'read', fraction: 1 })
  return { rows, columns }
}

/**
 * The worksheet XML, chunk by chunk. A 40 MB CSV becomes a worksheet several
 * times its size, and building that as one string is what made the import
 * need a gigabyte of memory; the chunks go into the archive as they come.
 */
export function* worksheetXmlChunks(
  rows: Iterable<readonly string[]>,
  shape: CsvGridShape,
  options?: { readonly onRow?: (rowsWritten: number) => void },
): Generator<string> {
  yield '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
  yield '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
  const lastColumn = columnLabel(Math.max(shape.columns, 1) - 1)
  yield `<dimension ref="A1:${lastColumn}${Math.max(shape.rows, 1)}"/>`
  yield '<sheetData>'
  let rowIndex = 0
  for (const row of rows) {
    const cells: string[] = []
    row.forEach((value, columnIndex) => {
      if (value === '') return
      const reference = `${columnLabel(columnIndex)}${rowIndex + 1}`
      cells.push(
        isNumericCell(value)
          ? `<c r="${reference}"><v>${Number(value)}</v></c>`
          : `<c r="${reference}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(encodeXlsxEscapes(value))}</t></is></c>`,
      )
    })
    if (cells.length > 0) yield `<row r="${rowIndex + 1}">${cells.join('')}</row>`
    rowIndex += 1
    options?.onRow?.(rowIndex)
  }
  yield '</sheetData></worksheet>'
}

export function buildWorksheetXml(rows: readonly (readonly string[])[]): string {
  return [...worksheetXmlChunks(rows, shapeOfRows(rows))].join('')
}

/**
 * Delimiter for the open-file path. The sniffer counts raw occurrences, so a
 * single-column prose file whose notes hold more semicolons/tabs than commas
 * ("hello; world") mis-sniffs and shatters into phantom columns — the same
 * trap sheetCsvToXlsxBuffer avoids by pinning comma for the AI grid path.
 * Guard: when the sniffed delimiter opens a single-field header row, the
 * body rows do not agree on a field count, and a comma parse stays
 * single-column throughout, the file is prose, not delimited data — keep it
 * one column. A genuine table keeps its delimiter: multi-field header, or a
 * title row over body rows that mostly share the same width (a comma-free
 * `;`/tab table must not collapse just because its first line is a title).
 * A `sep=` line is Excel's own declaration and skips the guard.
 */
export function resolveImportDelimiter(csvText: string): string {
  const declared = splitSepDeclaration(csvText).delimiter
  if (declared !== undefined) return declared
  const sniffed = sniffDelimiter(csvText)
  if (sniffed === ',') return sniffed
  const sniffedRows = parseCsv(csvText, sniffed)
  if (sniffedRows.length === 0 || sniffedRows[0]!.length !== 1) return sniffed
  if (bodyRowsAgreeOnWidth(sniffedRows.slice(1))) return sniffed
  if (parseCsv(csvText, ',').every((row) => row.length === 1)) return ','
  return sniffed
}

const TABLE_WIDTH_AGREEMENT = 0.6

function bodyRowsAgreeOnWidth(rows: readonly (readonly string[])[]): boolean {
  const widths = new Map<number, number>()
  for (const row of rows)
    if (row.length > 1) widths.set(row.length, (widths.get(row.length) ?? 0) + 1)
  const dominant = Math.max(0, ...widths.values())
  return dominant >= 2 && dominant >= rows.length * TABLE_WIDTH_AGREEMENT
}
