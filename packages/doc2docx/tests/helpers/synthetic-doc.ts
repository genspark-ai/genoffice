import * as CFB from 'cfb'

/**
 * Build minimal but structurally valid Word 97-2003 .doc files for the reader
 * tests: a FIB in the WordDocument stream, the piece text at a fixed offset
 * and a CLX piece table in the table stream. Every low-level byte layout is
 * overridable so the error paths can be exercised (bad FIB values, corrupt
 * CLX payloads, missing streams, ...).
 */

/** piece text starts after the FIB (1242 bytes) and is 4-byte aligned */
const TEXT_OFFSET = 2048
const TABLE_PAIR_COUNT = 136
const CLX_PAIR_INDEX = 33
const DEFAULT_WIDENT = 0xa5ec
const FLAG_ENCRYPTED = 0x0100
const FLAG_WHICH_TABLE_STREAM = 0x0200
const FLAG_OBFUSCATED = 0x8000
const FC_COMPRESSED = 0x40000000

export interface SyntheticPiece {
  text: string
  /** one byte per character (Windows-1252); default UTF-16LE */
  compressed?: boolean
}

export interface SyntheticPcd {
  /** byte offset of the piece text in the WordDocument stream */
  fc: number
  compressed?: boolean
  flags?: number
}

export interface SyntheticDocOptions {
  nFib?: number
  wIdent?: number
  /** extra FibBase.flags; fWhichTblStm / fEncrypted / fObfuscated are added from the options below */
  flags?: number
  encrypted?: boolean
  obfuscated?: boolean
  tableStream?: '0Table' | '1Table'
  /** body character count (FibRgLw97.ccpText); default = the pieces' total length */
  ccpText?: number
  pieces?: SyntheticPiece[]
  /** raw CLX bytes (overrides the piece table built from `pieces`) */
  clx?: Uint8Array
  /** raw WordDocument stream (overrides the generated FIB + text) */
  wordDocument?: Uint8Array
  /** raw table stream (overrides the generated CLX) */
  table?: Uint8Array
  /** pair-33 fc override (table-stream offset of the CLX) */
  fcClx?: number
  /** pair-33 lcb override */
  lcbClx?: number
  omitWordDocument?: boolean
  omitTableStream?: boolean
}

/** Windows-1252 high-half punctuation -> byte, for compressed-piece fixtures */
const CP1252_ENCODE: Record<string, number> = {
  '\u20ac': 0x80,
  '\u201a': 0x82,
  '\u0192': 0x83,
  '\u201e': 0x84,
  '\u2026': 0x85,
  '\u2020': 0x86,
  '\u2021': 0x87,
  '\u02c6': 0x88,
  '\u2030': 0x89,
  '\u0160': 0x8a,
  '\u2039': 0x8b,
  '\u0152': 0x8c,
  '\u017d': 0x8e,
  '\u2018': 0x91,
  '\u2019': 0x92,
  '\u201c': 0x93,
  '\u201d': 0x94,
  '\u2022': 0x95,
  '\u2013': 0x96,
  '\u2014': 0x97,
  '\u02dc': 0x98,
  '\u2122': 0x99,
  '\u0161': 0x9a,
  '\u203a': 0x9b,
  '\u0153': 0x9c,
  '\u017e': 0x9e,
  '\u0178': 0x9f,
}

function encodeCompressed(text: string): Uint8Array {
  const out = new Uint8Array(text.length)
  for (let index = 0; index < text.length; index++) {
    const code = text.charCodeAt(index)
    if (code < 0x80 || (code >= 0xa0 && code <= 0xff)) {
      out[index] = code
      continue
    }
    const byte = CP1252_ENCODE[text[index]!]
    if (byte === undefined)
      throw new Error(`test fixture: cannot encode ${JSON.stringify(text[index])} as Windows-1252`)
    out[index] = byte
  }
  return out
}

function encodeUtf16Le(text: string): Uint8Array {
  const out = new Uint8Array(text.length * 2)
  const view = new DataView(out.buffer)
  for (let index = 0; index < text.length; index++)
    view.setUint16(index * 2, text.charCodeAt(index), true)
  return out
}

/** the WordDocument bytes of one piece */
export function pieceBytes(piece: SyntheticPiece): Uint8Array {
  return piece.compressed ? encodeCompressed(piece.text) : encodeUtf16Le(piece.text)
}

/** a Pcdt CLX record from raw CPs and PCDs (fc = byte offset in the WordDocument stream) */
export function encodeClx(cps: number[], pcds: SyntheticPcd[]): Uint8Array {
  const lcb = 4 * cps.length + 8 * pcds.length
  const out = new Uint8Array(5 + lcb)
  const view = new DataView(out.buffer)
  out[0] = 0x02 // Pcdt
  view.setUint32(1, lcb, true)
  cps.forEach((cp, index) => view.setUint32(5 + index * 4, cp, true))
  pcds.forEach((pcd, index) => {
    const start = 5 + cps.length * 4 + index * 8
    view.setUint16(start, pcd.flags ?? 0, true)
    const raw = pcd.compressed ? ((pcd.fc * 2) | FC_COMPRESSED) >>> 0 : pcd.fc >>> 0
    view.setUint32(start + 2, raw, true)
    view.setUint16(start + 6, 0, true) // prm
  })
  return out
}

/** wrap streams in an OLE2 compound file (cfb seeds its own producer marker, ignored on read) */
export function buildCompound(streams: Array<[string, Uint8Array]>): Uint8Array {
  const container = CFB.utils.cfb_new()
  for (const [name, content] of streams) CFB.utils.cfb_add(container, name, Array.from(content))
  const written = CFB.write(container, { type: 'buffer', fileType: 'cfb' }) as Uint8Array
  return new Uint8Array(written)
}

interface FibValues {
  wIdent: number
  nFib: number
  flags: number
  ccpText: number
  fcClx: number
  lcbClx: number
}

function writeFib(target: Uint8Array, values: FibValues): void {
  const view = new DataView(target.buffer, target.byteOffset, target.byteLength)
  view.setUint16(0, values.wIdent, true)
  view.setUint16(2, values.nFib, true)
  view.setUint16(10, values.flags, true)
  let offset = 32
  view.setUint16(offset, 14, true) // csw
  offset += 2 + 14 * 2
  view.setUint16(offset, 22, true) // cslw
  offset += 2
  const lwStart = offset
  view.setUint32(lwStart, target.length, true) // cbMac
  view.setUint32(lwStart + 3 * 4, values.ccpText, true) // ccpText
  for (const index of [11, 14, 17]) view.setUint32(lwStart + index * 4, 0xffffff, true) // reserved values Word writes
  offset += 22 * 4
  view.setUint16(offset, TABLE_PAIR_COUNT, true) // cbRgFcLcb
  offset += 2
  view.setUint32(offset + CLX_PAIR_INDEX * 8, values.fcClx, true)
  view.setUint32(offset + CLX_PAIR_INDEX * 8 + 4, values.lcbClx, true)
}

/** build a .doc from FIB fields + text pieces (or raw stream overrides for corrupt cases) */
export function buildSyntheticDoc(options: SyntheticDocOptions = {}): Uint8Array {
  const tableStream = options.tableStream ?? '1Table'
  let flags = options.flags ?? 0
  if (tableStream === '1Table') flags |= FLAG_WHICH_TABLE_STREAM
  if (options.encrypted) flags |= FLAG_ENCRYPTED
  if (options.obfuscated) flags |= FLAG_OBFUSCATED
  const wIdent = options.wIdent ?? DEFAULT_WIDENT
  const nFib = options.nFib ?? 0x0101
  const pieces = options.pieces ?? []
  const fcClx = options.fcClx ?? 0

  // piece text + CP ranges, unless the caller supplies raw streams
  let clx = options.clx
  let wordDocument = options.wordDocument
  if (!wordDocument) {
    const cps = [0]
    const pcds: SyntheticPcd[] = []
    const encoded: Uint8Array[] = []
    let charCount = 0
    let byteCursor = TEXT_OFFSET
    for (const piece of pieces) {
      const bytes = pieceBytes(piece)
      pcds.push({ fc: byteCursor, compressed: piece.compressed === true })
      encoded.push(bytes)
      byteCursor += bytes.length
      charCount += piece.text.length
      cps.push(charCount)
    }
    clx ??= encodeClx(cps, pcds)
    const word = new Uint8Array(byteCursor)
    writeFib(word, {
      wIdent,
      nFib,
      flags,
      ccpText: options.ccpText ?? charCount,
      fcClx,
      lcbClx: options.lcbClx ?? clx.length,
    })
    let textOffset = TEXT_OFFSET
    for (const bytes of encoded) {
      word.set(bytes, textOffset)
      textOffset += bytes.length
    }
    wordDocument = word
  }
  let table = options.table
  if (!table) {
    table = new Uint8Array(fcClx + (clx?.length ?? 0))
    if (clx) table.set(clx, fcClx)
  }

  const streams: Array<[string, Uint8Array]> = []
  if (!options.omitWordDocument) streams.push(['WordDocument', wordDocument])
  if (!options.omitTableStream) streams.push([tableStream, table])
  return buildCompound(streams)
}
