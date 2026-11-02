import * as CFB from 'cfb'

/**
 * Build minimal but structurally valid Word 97-2003 .doc files for the reader
 * tests: a FIB in the WordDocument stream, the piece text at a fixed offset
 * and a CLX piece table in the table stream. Every low-level byte layout is
 * overridable so the error paths can be exercised (bad FIB values, corrupt
 * CLX payloads, missing streams, ...).
 */

/** piece text starts after the FIB (1242 bytes) and is 4-byte aligned */
export const SYNTHETIC_TEXT_OFFSET = 2048
const TEXT_OFFSET = SYNTHETIC_TEXT_OFFSET
/** FKP pages are 512-byte pages of the WordDocument stream */
export const SYNTHETIC_FKP_PAGE_BYTES = 512
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

export interface SyntheticChpxRun {
  fcStart: number
  fcEnd: number
  grpprl?: Uint8Array
}

export interface SyntheticPapxEntry {
  fcStart: number
  fcEnd: number
  istd?: number
  grpprl?: Uint8Array
  /** true -> BxPap.bOffset 0: the paragraph has no PAPX (default properties) */
  defaults?: boolean
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
  /** extra 512-byte CHPX FKP pages appended to the WordDocument stream (raw for corrupt cases) */
  chpxFkpPages?: Uint8Array[]
  /** extra 512-byte PAPX FKP pages appended to the WordDocument stream (raw for corrupt cases) */
  papxFkpPages?: Uint8Array[]
  /** raw PlcBteChpx override (appended to the table stream; takes precedence over chpxFkpPages) */
  chpxPlcf?: Uint8Array
  /** raw PlcBtePapx override (appended to the table stream; takes precedence over papxFkpPages) */
  papxPlcf?: Uint8Array
  /** pair-12/13 lcb overrides, for malformed-table tests */
  chpxPlcLcb?: number
  papxPlcLcb?: number
  /** extra Data stream (sprmPHugePapx property blocks) */
  dataStream?: Uint8Array
  /** non-body story character counts (FibRgLw97), for the stage-3a honest warnings */
  ccpFtn?: number
  ccpHdd?: number
  ccpAtn?: number
  ccpEdn?: number
  ccpTxbx?: number
  ccpHdrTxbx?: number
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

/**
 * A 512-byte CHPX FKP page from runs: rgfc (crun+1 FCs), rgb (offset/2 per run)
 * and the CHPX records packed after them; crun is the last byte. A run without
 * a grpprl gets rgb entry 0 (no explicit properties).
 */
export function encodeChpxFkp(runs: SyntheticChpxRun[]): Uint8Array {
  if (runs.length > 101) throw new Error('test fixture: too many CHPX runs for one FKP page')
  const page = new Uint8Array(SYNTHETIC_FKP_PAGE_BYTES)
  const view = new DataView(page.buffer)
  const rgbStart = 4 * (runs.length + 1)
  const fcs = runs.map((run) => run.fcStart)
  if (runs.length > 0) fcs.push(runs[runs.length - 1]!.fcEnd)
  fcs.forEach((fc, index) => view.setUint32(index * 4, fc >>> 0, true))
  let data = rgbStart + runs.length
  if (data % 2 !== 0) data++
  runs.forEach((run, index) => {
    const grpprl = run.grpprl ?? new Uint8Array(0)
    if (grpprl.length === 0) return
    if (data + 1 + grpprl.length > 511)
      throw new Error('test fixture: CHPX records overflow the FKP page')
    page[data] = grpprl.length
    page.set(grpprl, data + 1)
    page[rgbStart + index] = data / 2
    data += 1 + grpprl.length
    if (data % 2 !== 0) data++
  })
  page[511] = runs.length
  return page
}

/**
 * A 512-byte PAPX FKP page from paragraph entries: rgfc (paragraph start FCs +
 * the last paragraph end), BxPap records (bOffset + 12 reserved bytes) and the
 * PapxInFkp records packed after them. An istd + sprms block of odd length uses
 * the short form (cb = words), an even one the long form (cb 0 + word count).
 */
export function encodePapxFkp(entries: SyntheticPapxEntry[]): Uint8Array {
  if (entries.length > 29) throw new Error('test fixture: too many PAPX entries for one FKP page')
  const page = new Uint8Array(SYNTHETIC_FKP_PAGE_BYTES)
  const view = new DataView(page.buffer)
  const rgbxStart = 4 * (entries.length + 1)
  const fcs = entries.map((entry) => entry.fcStart)
  if (entries.length > 0) fcs.push(entries[entries.length - 1]!.fcEnd)
  fcs.forEach((fc, index) => view.setUint32(index * 4, fc >>> 0, true))
  let data = rgbxStart + entries.length * 13
  if (data % 2 !== 0) data++
  entries.forEach((entry, index) => {
    const offsetAt = rgbxStart + index * 13
    if (entry.defaults) {
      page[offsetAt] = 0
      return
    }
    const istd = entry.istd ?? 0
    const sprms = entry.grpprl ?? new Uint8Array(0)
    const raw = new Uint8Array(2 + sprms.length)
    raw[0] = istd & 0xff
    raw[1] = (istd >> 8) & 0xff
    raw.set(sprms, 2)
    let encoded: Uint8Array
    if (raw.length % 2 === 1) {
      const cb = (raw.length + 1) / 2
      if (cb > 255) throw new Error('test fixture: PAPX grpprl too long for the short form')
      encoded = new Uint8Array(1 + raw.length)
      encoded[0] = cb
      encoded.set(raw, 1)
    } else {
      const words = raw.length / 2
      if (words > 255) throw new Error('test fixture: PAPX grpprl too long for the long form')
      encoded = new Uint8Array(2 + raw.length)
      encoded[0] = 0
      encoded[1] = words
      encoded.set(raw, 2)
    }
    if (data + encoded.length > 511)
      throw new Error('test fixture: PAPX records overflow the FKP page')
    page.set(encoded, data)
    page[offsetAt] = data / 2
    data += encoded.length
    if (data % 2 !== 0) data++
  })
  page[511] = entries.length
  return page
}

/** a PlcBte table: n+1 FCs followed by n PnFkp page numbers */
export function encodePlcBte(fcs: number[], pages: number[]): Uint8Array {
  if (fcs.length !== pages.length + 1)
    throw new Error('test fixture: encodePlcBte needs pages.length + 1 FCs')
  const out = new Uint8Array(4 * (fcs.length + pages.length))
  const view = new DataView(out.buffer)
  fcs.forEach((fc, index) => view.setUint32(index * 4, fc >>> 0, true))
  const pagesStart = 4 * fcs.length
  pages.forEach((page, index) => view.setUint32(pagesStart + index * 4, page >>> 0, true))
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
  chpxPlcf?: { fc: number; lcb: number }
  papxPlcf?: { fc: number; lcb: number }
  ccpFtn?: number
  ccpHdd?: number
  ccpAtn?: number
  ccpEdn?: number
  ccpTxbx?: number
  ccpHdrTxbx?: number
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
  const storyCounts: Array<[number, number | undefined]> = [
    [4, values.ccpFtn], // ccpFtn
    [5, values.ccpHdd], // ccpHdd
    [7, values.ccpAtn], // ccpAtn
    [8, values.ccpEdn], // ccpEdn
    [9, values.ccpTxbx], // ccpTxbx
    [10, values.ccpHdrTxbx], // ccpHdrTxbx
  ]
  for (const [index, count] of storyCounts) {
    if (count !== undefined) view.setUint32(lwStart + index * 4, count, true)
  }
  for (const index of [11, 14, 17]) view.setUint32(lwStart + index * 4, 0xffffff, true) // reserved values Word writes
  offset += 22 * 4
  view.setUint16(offset, TABLE_PAIR_COUNT, true) // cbRgFcLcb
  offset += 2
  view.setUint32(offset + CLX_PAIR_INDEX * 8, values.fcClx, true)
  view.setUint32(offset + CLX_PAIR_INDEX * 8 + 4, values.lcbClx, true)
  const CHPX_PAIR_INDEX = 12
  const PAPX_PAIR_INDEX = 13
  if (values.chpxPlcf) {
    view.setUint32(offset + CHPX_PAIR_INDEX * 8, values.chpxPlcf.fc, true)
    view.setUint32(offset + CHPX_PAIR_INDEX * 8 + 4, values.chpxPlcf.lcb, true)
  }
  if (values.papxPlcf) {
    view.setUint32(offset + PAPX_PAIR_INDEX * 8, values.papxPlcf.fc, true)
    view.setUint32(offset + PAPX_PAIR_INDEX * 8 + 4, values.papxPlcf.lcb, true)
  }
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
  let chpxPlcf = options.chpxPlcf
  let papxPlcf = options.papxPlcf
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
    // FKP pages follow the text, one page per 512-byte block
    let nextPage = Math.ceil(byteCursor / SYNTHETIC_FKP_PAGE_BYTES)
    const chpxPages: number[] = []
    const papxPages: number[] = []
    for (const _page of options.chpxFkpPages ?? []) {
      chpxPages.push(nextPage)
      nextPage++
    }
    for (const _page of options.papxFkpPages ?? []) {
      papxPages.push(nextPage)
      nextPage++
    }
    chpxPlcf ??= chpxPages.length ? encodePlcBte([TEXT_OFFSET, ...chpxPages], chpxPages) : undefined
    papxPlcf ??= papxPages.length ? encodePlcBte([TEXT_OFFSET, ...papxPages], papxPages) : undefined
    // table layout: CLX, then the two formatting tables
    const baseTableLength = fcClx + clx.length
    const chpxPlcFc = baseTableLength
    const papxPlcFc = baseTableLength + (chpxPlcf?.length ?? 0)
    const word = new Uint8Array(nextPage * SYNTHETIC_FKP_PAGE_BYTES)
    writeFib(word, {
      wIdent,
      nFib,
      flags,
      ccpText: options.ccpText ?? charCount,
      fcClx,
      lcbClx: options.lcbClx ?? clx.length,
      ccpFtn: options.ccpFtn,
      ccpHdd: options.ccpHdd,
      ccpAtn: options.ccpAtn,
      ccpEdn: options.ccpEdn,
      ccpTxbx: options.ccpTxbx,
      ccpHdrTxbx: options.ccpHdrTxbx,
      ...(chpxPlcf
        ? { chpxPlcf: { fc: chpxPlcFc, lcb: options.chpxPlcLcb ?? chpxPlcf.length } }
        : {}),
      ...(papxPlcf
        ? { papxPlcf: { fc: papxPlcFc, lcb: options.papxPlcLcb ?? papxPlcf.length } }
        : {}),
    })
    let textOffset = TEXT_OFFSET
    for (const bytes of encoded) {
      word.set(bytes, textOffset)
      textOffset += bytes.length
    }
    ;[...(options.chpxFkpPages ?? []), ...(options.papxFkpPages ?? [])].forEach((page, index) => {
      const pageNumber = [...chpxPages, ...papxPages][index]!
      word.set(page, pageNumber * SYNTHETIC_FKP_PAGE_BYTES)
    })
    wordDocument = word
  }
  let table = options.table
  if (!table) {
    const baseTableLength = fcClx + (clx?.length ?? 0)
    table = new Uint8Array(baseTableLength + (chpxPlcf?.length ?? 0) + (papxPlcf?.length ?? 0))
    if (clx) table.set(clx, fcClx)
    if (chpxPlcf) table.set(chpxPlcf, baseTableLength)
    if (papxPlcf) table.set(papxPlcf, baseTableLength + (chpxPlcf?.length ?? 0))
  }

  const streams: Array<[string, Uint8Array]> = []
  if (!options.omitWordDocument) streams.push(['WordDocument', wordDocument])
  if (!options.omitTableStream) streams.push([tableStream, table])
  if (options.dataStream) streams.push(['Data', options.dataStream])
  return buildCompound(streams)
}
