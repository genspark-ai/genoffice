import { LegacyXlsError } from './errors'
import {
  BIFF,
  type BiffRecord,
  SegmentedReader,
  decodeRk,
  payloadOf,
  recordSegments,
  walkRecords,
} from './biff'

/** Caps a malformed or absurd workbook hits instead of exhausting the process. */
const MAX_SST_STRINGS = 1_000_000
const MAX_SHEETS = 4096
const MAX_CELLS = 20_000_000
const MAX_MERGES = 100_000

export type XlsCellType = 'string' | 'number' | 'boolean' | 'error' | 'blank'

export interface XlsCell {
  row: number
  col: number
  type: XlsCellType
  value: string | number | boolean | null
}

export interface XlsMerge {
  firstRow: number
  lastRow: number
  firstCol: number
  lastCol: number
}

export interface XlsSheet {
  name: string
  visible: boolean
  /** cells by row, then column; both zero-based like the records */
  rows: Map<number, Map<number, XlsCell>>
  dimension?: { firstRow: number; lastRow: number; firstCol: number; lastCol: number }
  merges: XlsMerge[]
  /** column widths in character units, keyed by zero-based column */
  colWidths: Map<number, number>
}

export interface XlsWorkbook {
  sheets: XlsSheet[]
  date1904: boolean
}

interface SheetRef {
  lbPlyPos: number
  name: string
  visible: boolean
}

const BIFF_ERROR_NAMES: Record<number, string> = {
  0x00: '#NULL!',
  0x07: '#DIV/0!',
  0x0f: '#VALUE!',
  0x17: '#REF!',
  0x1d: '#NAME?',
  0x24: '#NUM!',
  0x2a: '#N/A',
  0x2b: '#GETTING_DATA',
}

/** A fresh cell store that also refuses to grow past the cap. */
class CellStore {
  readonly rows = new Map<number, Map<number, XlsCell>>()
  private count = 0

  put(cell: XlsCell): void {
    if (this.count >= MAX_CELLS) {
      throw new LegacyXlsError('limits_exceeded', `more than ${MAX_CELLS} cells`)
    }
    let row = this.rows.get(cell.row)
    if (!row) {
      row = new Map()
      this.rows.set(cell.row, row)
    }
    if (!row.has(cell.col)) this.count += 1
    row.set(cell.col, cell)
  }
}

/** The shared string table: one segment-aware read across SST + CONTINUEs.
    Both counts are 4-byte fields (MS-XLS 2.4.265), not the 2-byte counts the
    old byte-builder wrote. */
function readSst(stream: Uint8Array, record: BiffRecord): string[] {
  const { segments } = recordSegments(stream, record)
  const reader = new SegmentedReader(segments)
  const total = reader.u32()
  const unique = reader.u32()
  void total
  const strings: string[] = []
  for (let i = 0; i < unique; i++) {
    if (strings.length >= MAX_SST_STRINGS) {
      throw new LegacyXlsError('limits_exceeded', `more than ${MAX_SST_STRINGS} strings`)
    }
    // XLUnicodeRichExtendedString (MS-XLS 2.5.293): cch, flags, [cRun],
    // [cbExtRst], chars, runs, ExtRst — the run/ext counts precede the text.
    const cch = reader.u16()
    const flags = reader.u8()
    const rich = (flags & 8) !== 0
    const farEast = (flags & 4) !== 0
    const runs = rich ? reader.u16() : 0
    const extSize = farEast ? reader.u32() : 0
    let text: string
    try {
      text = reader.chars(cch, (flags & 1) === 1)
    } catch (err) {
      throw new LegacyXlsError('bad_sst', err instanceof Error ? err.message : undefined)
    }
    if (runs > 0) reader.rawBytes(runs * 4)
    if (extSize > 0) reader.rawBytes(extSize)
    strings.push(text)
  }
  return strings
}

/** ShortXLUnicodeString (cch is ONE byte + flags + chars): sheet names and
    similar. `at` points at the cch byte inside the record payload. */
function readShortUnicodeString(view: DataView, at: number): { text: string; next: number } {
  const cch = view.getUint8(at)
  const flags = view.getUint8(at + 1)
  const high = (flags & 1) === 1
  const start = at + 2
  const size = high ? cch * 2 : cch
  const text = Buffer.from(view.buffer, view.byteOffset + start, size).toString(
    high ? 'utf16le' : 'latin1',
  )
  return { text, next: start + size }
}

/** Parse one worksheet substream: from its BOF to its EOF. */
function parseWorksheet(stream: Uint8Array, startRecord: BiffRecord, sst: string[]): XlsSheet {
  const sheet: XlsSheet = {
    name: '',
    visible: true,
    rows: new Map(),
    merges: [],
    colWidths: new Map(),
  }
  const store = new CellStore()
  let pendingStringCell: { row: number; col: number } | null = null
  let record: BiffRecord | undefined = startRecord
  let depth = 0
  while (record) {
    const payload = payloadOf(stream, record)
    const view = new DataView(payload.buffer, payload.byteOffset, payload.byteLength)
    switch (record.id) {
      case BIFF.bof:
        depth += 1
        break
      case BIFF.eof:
        depth -= 1
        if (depth === 0) {
          sheet.name = sheet.name || 'Sheet'
          sheet.rows = store.rows
          return sheet
        }
        break
      case BIFF.dimensions:
        sheet.dimension = {
          firstRow: view.getUint32(0, true),
          lastRow: view.getUint32(4, true) - 1,
          firstCol: view.getUint16(8, true),
          lastCol: view.getUint16(10, true) - 1,
        }
        break
      case BIFF.labelSst: {
        const row = view.getUint16(0, true)
        const col = view.getUint16(2, true)
        const isst = view.getUint32(6, true)
        if (isst >= sst.length) throw new LegacyXlsError('bad_sst', `isst ${isst} out of range`)
        store.put({ row, col, type: 'string', value: sst[isst]! })
        break
      }
      case BIFF.number: {
        const row = view.getUint16(0, true)
        const col = view.getUint16(2, true)
        store.put({ row, col, type: 'number', value: view.getFloat64(6, true) })
        break
      }
      case BIFF.rk: {
        const row = view.getUint16(0, true)
        const col = view.getUint16(2, true)
        store.put({ row, col, type: 'number', value: decodeRk(view.getInt32(6, true)) })
        break
      }
      case BIFF.mulRk: {
        const row = view.getUint16(0, true)
        const firstCol = view.getUint16(2, true)
        // count - 1 XF words then the final COL record (2 bytes)
        const rks = (payload.length - 6) / 6
        for (let i = 0; i < rks; i++) {
          const rk = view.getInt32(6 + i * 6, true)
          store.put({ row, col: firstCol + i, type: 'number', value: decodeRk(rk) })
        }
        break
      }
      case BIFF.boolErr: {
        const row = view.getUint16(0, true)
        const col = view.getUint16(2, true)
        const kind = view.getUint8(6)
        const value = view.getUint8(7)
        if (kind === 1) {
          store.put({ row, col, type: 'boolean', value: value === 1 })
        } else {
          store.put({
            row,
            col,
            type: 'error',
            value: BIFF_ERROR_NAMES[value] ?? '#VALUE!',
          })
        }
        break
      }
      case BIFF.blank:
        store.put({
          row: view.getUint16(0, true),
          col: view.getUint16(2, true),
          type: 'blank',
          value: null,
        })
        break
      case BIFF.mulBlank: {
        const row = view.getUint16(0, true)
        const firstCol = view.getUint16(2, true)
        const count = (payload.length - 6) / 2
        for (let i = 0; i < count; i++) {
          store.put({ row, col: firstCol + i, type: 'blank', value: null })
        }
        break
      }
      case BIFF.formula: {
        const row = view.getUint16(0, true)
        const col = view.getUint16(2, true)
        // FormulaValue (MS-XLS 2.5.133): the 8-byte result rides bytes 6..13;
        // fExprO (0xFFFF) is the LAST word, the value type is byte1 (offset 6)
        // and the bool/error payload is byte3 (offset 8).
        if (payload[12] === 0xff && payload[13] === 0xff) {
          const kind = payload[6]
          const value = payload[8]
          if (kind === 1) {
            store.put({ row, col, type: 'boolean', value: value === 1 })
          } else if (kind === 2) {
            store.put({ row, col, type: 'error', value: BIFF_ERROR_NAMES[value] ?? '#VALUE!' })
          } else {
            // string (0x00) and blank string (0x03): the text follows in the
            // String record immediately after this one
            pendingStringCell = { row, col }
            store.put({ row, col, type: 'string', value: '' })
          }
        } else {
          store.put({ row, col, type: 'number', value: view.getFloat64(6, true) })
        }
        break
      }
      case BIFF.string: {
        // the string result of the formula this substream just saw
        const text = readFormulaString(payload)
        if (pendingStringCell) {
          const row = store.rows.get(pendingStringCell.row)
          const cell = row?.get(pendingStringCell.col)
          if (cell?.type === 'string' && cell.value === '') {
            row!.set(pendingStringCell.col, { ...cell, value: text })
          }
          pendingStringCell = null
        }
        break
      }
      case BIFF.mergedCells: {
        const count = view.getUint16(0, true)
        if (count > MAX_MERGES) throw new LegacyXlsError('limits_exceeded', 'merge ranges')
        for (let i = 0; i < count; i++) {
          const at = 2 + i * 8
          // Ref8/Ref8U are zero-based (MS-XLS 2.5.209), so no bias correction.
          sheet.merges.push({
            firstRow: view.getUint16(at, true),
            lastRow: view.getUint16(at + 2, true),
            firstCol: view.getUint16(at + 4, true),
            lastCol: view.getUint16(at + 6, true),
          })
        }
        break
      }
      case BIFF.colInfo: {
        const first = view.getUint16(0, true)
        const last = view.getUint16(2, true)
        const widthChars = view.getUint16(4, true) / 256
        for (let c = first; c <= last; c++) sheet.colWidths.set(c, widthChars)
        break
      }
      default:
        break
    }
    record = walkFrom(stream, record).next().value
  }
  sheet.name = sheet.name || 'Sheet'
  return sheet
}

/** the formula string result: ShortXLUnicodeString-style cch + flags + chars */
function readFormulaString(payload: Uint8Array): string {
  const view = new DataView(payload.buffer, payload.byteOffset, payload.byteLength)
  const cch = view.getUint16(0, true)
  const high = (view.getUint8(2) & 1) === 1
  return Buffer.from(payload.buffer, payload.byteOffset + 3, high ? cch * 2 : cch).toString(
    high ? 'utf16le' : 'latin1',
  )
}

/** the next record after `record` */
function* walkFrom(stream: Uint8Array, record: BiffRecord): Generator<BiffRecord> {
  let at = record.offset + record.length
  while (at + 4 <= stream.length) {
    const id = stream[at]! | (stream[at + 1]! << 8)
    const length = stream[at + 2]! | (stream[at + 3]! << 8)
    const offset = at + 4
    if (offset + length > stream.length) {
      throw new LegacyXlsError('truncated', `record ${id.toString(16)} claims ${length} bytes`)
    }
    yield { id, length, offset }
    at = offset + length
  }
}

/** Parse the workbook: the globals substream, then every sheet's substream. */
export function parseXlsWorkbook(stream: Uint8Array): XlsWorkbook {
  const walker = walkRecords(stream)
  const first = walker.next()
  if (first.done) throw new LegacyXlsError('missing_stream')
  const bof = payloadOf(stream, first.value)
  if (first.value.id !== BIFF.bof || bof.length < 4) throw new LegacyXlsError('corrupt', 'no BOF')
  const version = bof[0]! | (bof[1]! << 8)
  const kind = bof[2]! | (bof[3]! << 8)
  if (kind !== 0x0005) throw new LegacyXlsError('corrupt', 'stream does not open with globals')
  if (version !== 0x0600)
    throw new LegacyXlsError('unsupported_version', `vers ${version.toString(16)}`)

  let date1904 = false
  const sheetRefs: SheetRef[] = []
  let sst: string[] = []
  let record: BiffRecord | undefined = first.value
  while (record) {
    switch (record.id) {
      case BIFF.eof:
        record = undefined
        continue
      case BIFF.filePass:
        throw new LegacyXlsError('encrypted')
      case BIFF.sst:
        sst = readSst(stream, record)
        break
      case BIFF.boundSheet: {
        if (sheetRefs.length >= MAX_SHEETS) {
          throw new LegacyXlsError('limits_exceeded', `more than ${MAX_SHEETS} sheets`)
        }
        const payload = payloadOf(stream, record)
        const pv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength)
        const lbPlyPos = pv.getUint32(0, true)
        const grbit = pv.getUint16(4, true)
        const state = grbit & 0x0003
        // lbPlyPos(4) + grbit(2) = the name's cch byte sits at payload offset 6
        const { text } = readShortUnicodeString(pv, 6)
        sheetRefs.push({ lbPlyPos, name: text, visible: state === 0 })
        break
      }
      case BIFF.date1904: {
        // Date1904 is always present in BIFF8 (MS-XLS 2.4.74); its 2-byte
        // payload carries f1904, so presence alone does not mean the 1904 epoch.
        const payload = payloadOf(stream, record)
        const pv = new DataView(payload.buffer, payload.byteOffset, payload.byteLength)
        date1904 = pv.byteLength >= 2 && (pv.getUint16(0, true) & 1) === 1
        break
      }
      default:
        break
    }
    if (record === undefined) break
    record = walker.next().value ?? undefined
  }

  const sheets: XlsSheet[] = sheetRefs.map((ref) => {
    // each sheet's substream opens at its own BOF, reached by seeking the stream
    // lbPlyPos points at the substream's record header; walkRecords yields the
    // payload offset, which sits 4 bytes in
    const seeker = walkRecords(stream)
    let target = seeker.next()
    while (!target.done && target.value.offset - 4 !== ref.lbPlyPos) target = seeker.next()
    if (target.done) throw new LegacyXlsError('bad_boundsheet', `sheet "${ref.name}"`)
    const sheet = parseWorksheet(stream, target.value, sst)
    sheet.name = ref.name
    sheet.visible = ref.visible
    return sheet
  })
  return { sheets, date1904 }
}
