import * as CFB from 'cfb'

/**
 * A minimal BIFF8 workbook builder: records written by hand, so the tests
 * exercise the reader against bytes whose content is known exactly. Records
 * follow MS-XLS: a 2-byte id, a 2-byte length, the payload.
 */

export type CellSpec =
  | { kind: 'label'; row: number; col: number; stringIndex: number }
  | { kind: 'number'; row: number; col: number; value: number }
  | { kind: 'rk'; row: number; col: number; rk: number }
  | { kind: 'bool'; row: number; col: number; value: boolean }
  | { kind: 'error'; row: number; col: number; code: number }
  | { kind: 'formulaNumber'; row: number; col: number; value: number }
  | { kind: 'formulaString'; row: number; col: number; text: string }

export interface SheetSpec {
  name: string
  hidden?: boolean
  dimensions?: { lastRow: number; lastCol: number }
  merges?: Array<[number, number, number, number]>
  colWidths?: Array<{ from: number; to: number; width: number }>
  cells: CellSpec[]
}

const record = (id: number, payload: Uint8Array): Uint8Array => {
  const out = new Uint8Array(4 + payload.length)
  const v = new DataView(out.buffer)
  v.setUint16(0, id, true)
  v.setUint16(2, payload.length, true)
  out.set(payload, 4)
  return out
}

const u8 = (value: number): Uint8Array => new Uint8Array([value & 0xff])

const u16 = (value: number): Uint8Array => {
  const out = new Uint8Array(2)
  new DataView(out.buffer).setUint16(0, value, true)
  return out
}

const u32 = (value: number): Uint8Array => {
  const out = new Uint8Array(4)
  new DataView(out.buffer).setUint32(0, value >>> 0, true)
  return out
}

const f64 = (value: number): Uint8Array => {
  const out = new Uint8Array(8)
  new DataView(out.buffer).setFloat64(0, value, true)
  return out
}

const concat = (parts: Uint8Array[]): Uint8Array => {
  const total = parts.reduce((sum, p) => sum + p.length, 0)
  const out = new Uint8Array(total)
  let at = 0
  for (const part of parts) {
    out.set(part, at)
    at += part.length
  }
  return out
}

const isHigh = (text: string): boolean => [...text].some((c) => c.charCodeAt(0) > 0xff)

const stringChars = (text: string): Uint8Array =>
  new Uint8Array(isHigh(text) ? Buffer.from(text, 'utf16le') : Buffer.from(text, 'latin1'))

/** XLUnicodeString body: cch + flags + chars */
const xlUnicodeString = (text: string): Uint8Array =>
  concat([u16(text.length), u8(isHigh(text) ? 1 : 0), stringChars(text)])

const bof = (kind: number): Uint8Array =>
  record(0x0809, concat([u16(0x0600), u16(kind), u16(0), u16(0), u32(0), u32(0)]))

const eof = (): Uint8Array => record(0x000a, new Uint8Array(0))

/** The SST as one record (no CONTINUEs). */
export function buildSstRecord(strings: string[]): Uint8Array {
  return record(
    0x00fc,
    concat([u16(strings.length), u16(strings.length), ...strings.map(xlUnicodeString)]),
  )
}

/** The shared string table split across SST + CONTINUEs at whole-string
    boundaries, each CONTINUE restarting the compression flag. */
export function buildSstSplit(strings: string[], perSegment: number): Uint8Array {
  const parts: Uint8Array[] = []
  let at = 0
  let first = true
  while (at < strings.length) {
    const end = Math.min(at + (first ? perSegment : 1), strings.length)
    const payload: Uint8Array[] = []
    if (first) payload.push(u16(strings.length), u16(strings.length))
    // a CONTINUE that lands between strings carries no flag byte: each
    // string's own header governs its characters
    for (let i = at; i < end; i++) payload.push(xlUnicodeString(strings[i]!))
    parts.push(record(first ? 0x00fc : 0x003c, concat(payload)))
    first = false
    at = end
  }
  return concat(parts)
}

/** Build one worksheet substream (BOF … EOF). */
export function buildSheetStream(sheet: SheetSpec): Uint8Array {
  const body: Uint8Array[] = []
  if (sheet.dimensions) {
    body.push(
      record(
        0x0200,
        concat([
          u32(0),
          u32(sheet.dimensions.lastRow + 1),
          u16(0),
          u16(sheet.dimensions.lastCol + 1),
          u16(0),
        ]),
      ),
    )
  }
  for (const cell of sheet.cells) {
    switch (cell.kind) {
      case 'label':
        body.push(
          record(0x00fd, concat([u16(cell.row), u16(cell.col), u16(0), u32(cell.stringIndex)])),
        )
        break
      case 'number':
        body.push(record(0x0203, concat([u16(cell.row), u16(cell.col), u16(0), f64(cell.value)])))
        break
      case 'rk':
        body.push(record(0x027e, concat([u16(cell.row), u16(cell.col), u16(0), u32(cell.rk)])))
        break
      case 'bool':
        body.push(
          record(
            0x0205,
            concat([u16(cell.row), u16(cell.col), u16(0), u8(1), u8(cell.value ? 1 : 0)]),
          ),
        )
        break
      case 'error':
        body.push(
          record(0x0205, concat([u16(cell.row), u16(cell.col), u16(0), u8(2), u8(cell.code)])),
        )
        break
      case 'formulaNumber':
        body.push(
          record(
            0x0006,
            concat([u16(cell.row), u16(cell.col), u16(0), f64(cell.value), u16(0), u32(0), u16(0)]),
          ),
        )
        break
      case 'formulaString':
        body.push(
          record(
            0x0006,
            concat([
              u16(cell.row),
              u16(cell.col),
              u16(0),
              new Uint8Array([0xff, 0xff, 0x00, 0x00, 0, 0, 0, 0]),
              u16(0),
              u32(0),
              u16(0),
            ]),
          ),
          record(0x0207, xlUnicodeString(cell.text)),
        )
        break
    }
  }
  for (const w of sheet.colWidths ?? []) {
    body.push(
      record(0x007d, concat([u16(w.from), u16(w.to), u16(Math.round(w.width * 256)), u16(0)])),
    )
  }
  if (sheet.merges?.length) {
    const entries = sheet.merges.map(([r1, r2, c1, c2]) =>
      concat([u16(r1 + 1), u16(r2 + 1), u16(c1 + 1), u16(c2 + 1)]),
    )
    body.push(record(0x00e5, concat([u16(sheet.merges.length), ...entries])))
  }
  return concat([bof(0x0010), ...body, eof()])
}

/** The workbook stream: globals (bound sheets + SST) then each sheet's substream. */
export function buildWorkbookStream(
  sheets: SheetSpec[],
  strings: string[],
  sstRecord?: Uint8Array,
): Uint8Array {
  const sst = sstRecord ?? buildSstRecord(strings)
  // bound-sheet records have a fixed shape per name, so their size is known
  // before the offsets they carry exist: size = header(4) + lbPlyPos(4) + grbit(2)
  //   + cch(1) + flags(1) + chars
  const boundSize = (name: string): number => 12 + stringChars(name).length
  const globalsLen =
    sheets.reduce((sum, s) => sum + boundSize(s.name), 0) + sst.length + eof().length
  // the globals BOF record is a full 20 bytes on the wire (header + 16 payload)
  let offset = 20 + globalsLen
  const boundRecords: Uint8Array[] = []
  for (const sheet of sheets) {
    boundRecords.push(
      record(
        0x0085,
        concat([
          u32(offset),
          u16(sheet.hidden ? 1 : 0),
          u8(sheet.name.length),
          u8(isHigh(sheet.name) ? 1 : 0),
          stringChars(sheet.name),
        ]),
      ),
    )
    offset += buildSheetStream(sheet).length
  }
  return concat([bof(0x0005), ...boundRecords, sst, eof(), ...sheets.map(buildSheetStream)])
}

/** The stream inside a real OLE2/CFB container, the way Excel writes it. */
export function buildXlsContainer(stream: Uint8Array): Uint8Array {
  const cfb = CFB.utils.cfb_new({ root: 'Root Entry' })
  CFB.utils.cfb_add(cfb, '/Workbook', Buffer.from(stream))
  const out = CFB.write(cfb, { type: 'buffer' }) as Buffer
  return new Uint8Array(out.buffer, out.byteOffset, out.byteLength)
}

export { u16, u32, u8, record, concat, xlUnicodeString }
