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
  /** a stylesheet built from synthetic styles (appended to the table stream, FIB pair 1) */
  stshStyles?: SyntheticStyle[]
  /** STD base length of the synthetic stylesheet (0x0A Word 97, 0x12 Word 2000+) */
  stshCbStdBase?: 0x0a | 0x12
  /** a raw STSH override (for corrupt-stylesheet tests) */
  stsh?: Uint8Array
  /** pair-1 lcb override, for malformed stylesheet tests */
  lcbStshf?: number
  /** section table entries (PlcfSed + SEPX); omitted -> the FIB carries no section table */
  sections?: SyntheticSection[]
  /** six header/footer stories of the first/only section (evenH, oddH, evenF, oddF, firstH, firstF) */
  headerStories?: Array<string | undefined>
  /** per-section header/footer stories (one six-entry array per section; overrides headerStories) */
  sectionHeaderStories?: Array<Array<string | undefined>>
  /** raw PlcfSed override + pair-6 lcb override (corrupt-section tests) */
  plcfSed?: Uint8Array
  lcbPlcfsed?: number
  /** raw PlcfHdd override + pair-11 lcb override (corrupt-header tests) */
  plcfHdd?: Uint8Array
  lcbPlcfhdd?: number
}

/** One synthetic section for the PlcfSed/SEPX builders. */
export interface SyntheticSection {
  /** section text end CP; defaults to the whole text for the last section */
  endCp?: number
  /** write fcSepx = 0xFFFFFFFF (Word's default section) */
  defaultSep?: boolean
  /** section sprms for the SEPX */
  sepxSprms?: Uint8Array
  /** force the SED's fcSepx, for out-of-range SEPX tests */
  fcSepxOverride?: number
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

export type SyntheticStyleKind = 'paragraph' | 'character' | 'table' | 'numbering'

/** One style for encodeStsh: array position = istd. */
export interface SyntheticStyle {
  /** built-in style id (default 4094 = user-defined) */
  sti?: number
  kind?: SyntheticStyleKind
  basedOn?: number
  next?: number
  name?: string
  /** PAPX sprms WITHOUT the 2-byte istd prefix (paragraph styles) */
  papxSprms?: Uint8Array
  /** CHPX sprms (character styles, and paragraph styles' UPX[1]) */
  chpxSprms?: Uint8Array
  /** write a cbStd == 0 empty slot */
  empty?: boolean
}

const STK_BY_KIND: Record<SyntheticStyleKind, number> = {
  paragraph: 1,
  character: 2,
  table: 3,
  numbering: 4,
}

function u16le(value: number): number[] {
  return [value & 0xff, (value >> 8) & 0xff]
}

/**
 * Build an STSH stylesheet (Word 97 layout): STSHI + one STD per array slot.
 * A paragraph style stores its PAPX grpprl (with the 2-byte istd prefix) as
 * UPX[0] and its CHPX as UPX[1]; a character style stores its CHPX as UPX[0].
 * cbStdBase 0x12 writes the extra 8-byte StdfPost2000 block Word 2000+ uses.
 */
export function encodeStsh(styles: SyntheticStyle[], cbStdBase: 0x0a | 0x12 = 0x0a): Uint8Array {
  const bytes: number[] = []
  const pushU16 = (value: number): void => {
    bytes.push(...u16le(value))
  }
  const stshi: number[] = [
    ...u16le(styles.length), // cstd
    ...u16le(cbStdBase), // cbSTDBaseInFile
    ...u16le(1), // info3
    ...u16le(0x0fff), // stiMaxWhenSaved
    ...u16le(0x000f), // istdMaxFixedWhenSaved
    ...u16le(0), // nVerBuiltInNamesWhenSaved
    ...u16le(0), // ftcAsci
    ...u16le(0), // ftcFE
    ...u16le(0), // ftcOther
  ]
  pushU16(stshi.length)
  bytes.push(...stshi)
  styles.forEach((style, istd) => {
    if (style.empty) {
      pushU16(0)
      return
    }
    const kind = style.kind ?? 'paragraph'
    const upx: number[][] = []
    if (kind === 'paragraph') {
      const papx = Uint8Array.from([...u16le(istd), ...(style.papxSprms ?? new Uint8Array(0))])
      upx.push([...papx])
      if (style.chpxSprms && style.chpxSprms.length > 0) upx.push([...style.chpxSprms])
    } else if (kind === 'character') {
      upx.push([...(style.chpxSprms ?? new Uint8Array(0))])
    }
    const payload: number[] = []
    const pushPayload = (value: number): void => {
      payload.push(...u16le(value))
    }
    pushPayload(style.sti ?? 4094) // info1: sti
    pushPayload(STK_BY_KIND[kind] | (((style.basedOn ?? 0x0fff) & 0x0fff) << 4)) // info2
    pushPayload(upx.length | (((style.next ?? 0x0fff) & 0x0fff) << 4)) // info3
    pushPayload(0) // bchUpe
    pushPayload(0) // grfstd
    if (cbStdBase === 0x12) payload.push(0, 0, 0, 0, 0, 0, 0, 0) // StdfPost2000
    const name = style.name ?? ''
    pushPayload(name.length)
    for (const char of name) pushPayload(char.charCodeAt(0))
    pushPayload(0) // null terminator
    for (const entry of upx) {
      pushPayload(entry.length)
      payload.push(...entry)
      if (entry.length % 2 === 1) payload.push(0)
    }
    pushU16(payload.length)
    bytes.push(...payload)
    if (payload.length % 2 === 1) bytes.push(0)
  })
  return Uint8Array.from(bytes)
}

/** a SEPX blob: u16 length + section sprms */
export function encodeSepx(sprms: Uint8Array): Uint8Array {
  const out = new Uint8Array(2 + sprms.length)
  new DataView(out.buffer).setUint16(0, sprms.length, true)
  out.set(sprms, 2)
  return out
}

/**
 * A PlcfSed table: n+1 section CPs then n 12-byte SEDs (fn u16, fcSepx u32,
 * fnMpr u16, fcMpr u32). undefined/0xFFFFFFFF fcSepx means the default SEP.
 */
export function encodeSectionPlcf(
  endCps: number[],
  sepxFcs: Array<number | undefined>,
): Uint8Array {
  const count = endCps.length
  if (sepxFcs.length !== count)
    throw new Error('test fixture: encodeSectionPlcf needs one fcSepx per section')
  const out = new Uint8Array(4 * (count + 1) + 12 * count)
  const view = new DataView(out.buffer)
  view.setUint32(0, 0, true)
  endCps.forEach((cp, index) => view.setUint32(4 * (index + 1), cp, true))
  sepxFcs.forEach((fc, index) => {
    const sedStart = 4 * (count + 1) + 12 * index
    view.setUint16(sedStart, 0, true) // fn
    view.setUint32(sedStart + 2, fc === undefined ? 0xffffffff : fc >>> 0, true) // fcSepx
    view.setUint16(sedStart + 6, 0, true) // fnMpr
    view.setUint32(sedStart + 8, 0, true) // fcMpr
  })
  return out
}

/** a PlcfHdd table: the story CPs only (no data entries) */
export function encodePlcfHdd(cps: number[]): Uint8Array {
  const out = new Uint8Array(4 * cps.length)
  const view = new DataView(out.buffer)
  cps.forEach((cp, index) => view.setUint32(index * 4, cp, true))
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
  stshf?: { fc: number; lcb: number }
  plcfsed?: { fc: number; lcb: number }
  plcfhdd?: { fc: number; lcb: number }
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
  const STSHF_PAIR_INDEX = 1
  const PLCFSED_PAIR_INDEX = 6
  const PLCFHDD_PAIR_INDEX = 11
  const CHPX_PAIR_INDEX = 12
  const PAPX_PAIR_INDEX = 13
  if (values.stshf) {
    view.setUint32(offset + STSHF_PAIR_INDEX * 8, values.stshf.fc, true)
    view.setUint32(offset + STSHF_PAIR_INDEX * 8 + 4, values.stshf.lcb, true)
  }
  for (const [index, pair] of [
    [PLCFSED_PAIR_INDEX, values.plcfsed],
    [PLCFHDD_PAIR_INDEX, values.plcfhdd],
  ] as const) {
    if (!pair) continue
    view.setUint32(offset + index * 8, pair.fc, true)
    view.setUint32(offset + index * 8 + 4, pair.lcb, true)
  }
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
  let plcfsed = options.plcfSed
  let plcfhdd = options.plcfHdd
  const stsh =
    options.stsh ??
    (options.stshStyles ? encodeStsh(options.stshStyles, options.stshCbStdBase) : undefined)
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
    const bodyCharCount = charCount
    // header/footer stories per section (six each) and the header-document text
    const sectionStorySets =
      options.sectionHeaderStories ?? (options.headerStories ? [options.headerStories] : [])
    const sectionDefs: SyntheticSection[] =
      options.sections ??
      (sectionStorySets.length > 0
        ? sectionStorySets.map(() => ({}))
        : options.headerStories
          ? [{}]
          : [])
    const storySets = sectionDefs.map(
      (_, index) =>
        sectionStorySets[index] ??
        (index === 0 && options.headerStories ? options.headerStories : []),
    )
    const storiesProvided =
      options.headerStories !== undefined || options.sectionHeaderStories !== undefined
    let headerText = ''
    let headerChars = 0
    const hddCps: number[] = []
    if (sectionDefs.length > 0) {
      if ((options.ccpFtn ?? 0) > 0)
        throw new Error(
          'test fixture: the builder does not support footnotes together with headers',
        )
      hddCps.push(0, 0, 0, 0, 0, 0, 0) // the six empty footnote/endnote separator stories
      for (const stories of storySets) {
        for (let variant = 0; variant < 6; variant++) {
          const text = stories[variant]
          if (text) {
            headerText += text + '\r'
            headerChars += text.length + 1
          }
          hddCps.push(headerChars)
        }
      }
      hddCps.push(headerChars + 1) // n+2nd CP (one past the last story limit)
    }
    if (headerText !== '') {
      pcds.push({ fc: byteCursor, compressed: false })
      const bytes = encodeUtf16Le(headerText)
      encoded.push(bytes)
      byteCursor += bytes.length
      charCount += headerText.length
      cps.push(charCount)
    }
    clx ??= encodeClx(cps, pcds)
    const totalTextLength = charCount
    // section CP ranges + SEPX placement after the FKP pages
    const sectionCps: number[] = []
    for (let index = 0; index < sectionDefs.length; index++) {
      const definition = sectionDefs[index]!
      let endCp = definition.endCp
      if (endCp === undefined) {
        if (index !== sectionDefs.length - 1)
          throw new Error('test fixture: sections other than the last need an endCp')
        endCp = totalTextLength
      }
      if (sectionCps.length > 0 && endCp < sectionCps[sectionCps.length - 1]!)
        throw new Error('test fixture: section end CPs must increase')
      sectionCps.push(endCp)
    }
    const sepxBlobs = sectionDefs.map((definition) =>
      definition.defaultSep || definition.fcSepxOverride !== undefined
        ? new Uint8Array(0)
        : encodeSepx(definition.sepxSprms ?? new Uint8Array(0)),
    )
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
    const sepxStart = nextPage * SYNTHETIC_FKP_PAGE_BYTES
    let sepxCursor = sepxStart
    const sepxFcs: Array<number | undefined> = sectionDefs.map((definition, index) => {
      if (definition.fcSepxOverride !== undefined) return definition.fcSepxOverride
      if (definition.defaultSep) return undefined
      const fc = sepxCursor
      sepxCursor += sepxBlobs[index]!.length
      return fc
    })
    const sepxTotal = sepxCursor - sepxStart
    plcfsed ??= sectionDefs.length > 0 ? encodeSectionPlcf(sectionCps, sepxFcs) : undefined
    plcfhdd ??=
      sectionDefs.length > 0 && (storiesProvided || headerText !== '')
        ? encodePlcfHdd(hddCps)
        : undefined
    // table layout: CLX, formatting tables, stylesheet, section and header tables
    const baseTableLength = fcClx + clx.length
    const chpxPlcFc = baseTableLength
    const papxPlcFc = chpxPlcFc + (chpxPlcf?.length ?? 0)
    const stshFc = papxPlcFc + (papxPlcf?.length ?? 0)
    const plcfsedFc = stshFc + (stsh?.length ?? 0)
    const plcfhddFc = plcfsedFc + (plcfsed?.length ?? 0)
    const word = new Uint8Array(nextPage * SYNTHETIC_FKP_PAGE_BYTES + sepxTotal)
    writeFib(word, {
      wIdent,
      nFib,
      flags,
      ccpText: options.ccpText ?? bodyCharCount,
      fcClx,
      lcbClx: options.lcbClx ?? clx.length,
      ccpFtn: options.ccpFtn,
      ccpHdd: options.ccpHdd ?? headerChars,
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
      ...(stsh ? { stshf: { fc: stshFc, lcb: options.lcbStshf ?? stsh.length } } : {}),
      ...(plcfsed ? { plcfsed: { fc: plcfsedFc, lcb: options.lcbPlcfsed ?? plcfsed.length } } : {}),
      ...(plcfhdd ? { plcfhdd: { fc: plcfhddFc, lcb: options.lcbPlcfhdd ?? plcfhdd.length } } : {}),
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
    sectionDefs.forEach((definition, index) => {
      const fc = sepxFcs[index]
      if (fc === undefined || definition.defaultSep || sepxBlobs[index]!.length === 0) return
      word.set(sepxBlobs[index]!, fc)
    })
    wordDocument = word
  }
  let table = options.table
  if (!table) {
    const baseTableLength = fcClx + (clx?.length ?? 0)
    table = new Uint8Array(
      baseTableLength +
        (chpxPlcf?.length ?? 0) +
        (papxPlcf?.length ?? 0) +
        (stsh?.length ?? 0) +
        (plcfsed?.length ?? 0) +
        (plcfhdd?.length ?? 0),
    )
    if (clx) table.set(clx, fcClx)
    if (chpxPlcf) table.set(chpxPlcf, baseTableLength)
    let cursor = baseTableLength + (chpxPlcf?.length ?? 0)
    if (papxPlcf) {
      table.set(papxPlcf, cursor)
      cursor += papxPlcf.length
    }
    if (stsh) {
      table.set(stsh, cursor)
      cursor += stsh.length
    }
    if (plcfsed) {
      table.set(plcfsed, cursor)
      cursor += plcfsed.length
    }
    if (plcfhdd) table.set(plcfhdd, cursor)
  }

  const streams: Array<[string, Uint8Array]> = []
  if (!options.omitWordDocument) streams.push(['WordDocument', wordDocument])
  if (!options.omitTableStream) streams.push([tableStream, table])
  if (options.dataStream) streams.push(['Data', options.dataStream])
  return buildCompound(streams)
}
