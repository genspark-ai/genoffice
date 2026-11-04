import { deflateSync } from 'node:zlib'
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
  /** per-section header/footer stories (one six-entry array per section; overrides headerStories); '' = blank, undefined = inherited */
  sectionHeaderStories?: Array<Array<string | undefined>>
  /** raw PlcfSed override + pair-6 lcb override (corrupt-section tests) */
  plcfSed?: Uint8Array
  lcbPlcfsed?: number
  /** raw PlcfHdd override + pair-11 lcb override (corrupt-header tests) */
  plcfHdd?: Uint8Array
  lcbPlcfhdd?: number
  /** footnotes: the body CP of the reference mark plus the note's own text (paragraph marks added by the builder) */
  footnotes?: SyntheticNote[]
  endnotes?: SyntheticNote[]
  /** write the note text but no PlcffndTxt / PlcfendTxt (degraded-document tests) */
  omitNoteTextTables?: boolean
  /** pair-3 lcb override (corrupt footnote text table tests) */
  lcbPlcffndTxt?: number
  /** synthetic PlfLst lists (table stream, FIB pair 73) + PlfLfo overrides (pair 74) */
  lists?: SyntheticListDefinition[]
  /** 1-based ilfo -> lsid overrides; defaults to one override per list in order */
  listOverrides?: number[]
  /** LFOLVL level overrides per ilfo (same order as listOverrides) */
  listLevelOverrides?: Array<SyntheticLevelOverride[] | undefined>
  /** raw PlfLst / PlfLfo overrides + pair-73/74 lcb overrides (corrupt-list tests) */
  plfLst?: Uint8Array
  lcbPlfLst?: number
  plfLfo?: Uint8Array
  lcbPlfLfo?: number
  /** font table names (SttbfFfn, FIB pair 15); array position = ftc */
  fonts?: string[]
  /** raw SttbfFfn override + pair-15 lcb override (corrupt-font-table tests) */
  sttbfFfn?: Uint8Array
  lcbSttbfFfn?: number
  /** STSHI rgftcStandardChpStsh of the synthetic stylesheet (ftcAsci, ftcFE, ftcOther) */
  stshDefaultFonts?: [number, number, number]
  /** floating shapes: PlcSpaMom/PlcSpaHdr anchors + DggInfo OfficeArt (blips in the WordDocument stream) */
  floatingShapes?: SyntheticFloatingShape[]
  /** comments: stories appended after the header text, PlcfandRef/PlcfandTxt, authors and annotation bookmarks */
  comments?: SyntheticComment[]
  /** leave FIB pair 5 empty although ccpAtn > 0 (degradation test) */
  omitPlcfandTxt?: boolean
  /** pair-4 lcb override (corrupt comment table test) */
  lcbPlcfandRef?: number
  /** text box stories (PlcftxbxTxt / PlcfHdrtxbxTxt) appended after the header text, keyed by shape id */
  textBoxStories?: SyntheticTextBoxStory[]
}

export interface SyntheticComment {
  /** comment paragraphs (the builder prepends the annotation reference character) */
  text: string | string[]
  author: string
  initials: string
  /** body CP of the \x05 annotation reference character (the test's body text must carry it) */
  referenceCp: number
  /** commented body range [start, end); omitted = point comment */
  range?: [number, number]
  /** ATRDPost10 dttm (Word 2002+ date); omitted = no AtrdExtra table */
  dttm?: number
}

export interface SyntheticTextBoxStory {
  story?: 'body' | 'header'
  spid: number
  text: string
  /** FTXBXS.fReusable: a deleted slot that carries no story */
  reusable?: boolean
}

export type SyntheticBlipKind = 'png' | 'jpeg' | 'wmf' | 'emf' | 'dib' | 'pict'

export interface SyntheticBlip {
  kind: SyntheticBlipKind
  bytes: Uint8Array
  /** metafiles: store deflate-compressed (the default) or raw */
  uncompressed?: boolean
  /** declared record length override (truncation tests) */
  recordLength?: number
  /** put the blip right after its FBSE inside the blip store instead of the WordDocument stream */
  embedded?: boolean
}

export interface SyntheticShapeChild {
  /** child anchor in the group's coordinate space */
  anchor: [number, number, number, number]
  blip?: SyntheticBlip
  group?: { bounds: [number, number, number, number]; children: SyntheticShapeChild[] }
}

export interface SyntheticFloatingShape {
  story: 'body' | 'header'
  cp: number
  spid?: number
  /** xaLeft, yaTop, xaRight, yaBottom in twips */
  rect: [number, number, number, number]
  bx?: number
  by?: number
  wr?: number
  belowText?: boolean
  /** a picture shape (omit for a non-picture shape such as a text box) */
  blip?: SyntheticBlip
  /** a group shape whose members are placed through its coordinate bounds */
  group?: { bounds: [number, number, number, number]; children: SyntheticShapeChild[] }
  /** pib override (missing-blip tests) */
  pib?: number
  /** extra OfficeArt properties ([pid, value]) of a non-picture shape */
  props?: Array<[number, number]>
}

/** One synthetic list level (encoded as LVLF + papx + chpx + Xst). */
export interface SyntheticListLevel {
  startAt?: number
  /** Word nfc numbering format (0 decimal, 23 bullet, ...) */
  numberFormat?: number
  /** raw Xst characters; 0x00-0x08 are level-number placeholders */
  levelText?: string
  /** jc bits of LVLF.info (0 left, 1 center, 2 right) */
  alignment?: number
  legal?: boolean
  noRestart?: boolean
  /** ixchFollow: 0 tab, 1 space, 2 nothing */
  follow?: number
  indentLeft?: number
  /** positive = first line, negative = hanging */
  indentFirstLine?: number
  indentRight?: number
  /** raw grpprlPapx override (corrupt-level tests) */
  paragraphSprms?: Uint8Array
  /** raw grpprlChpx (level number run properties) */
  characterSprms?: Uint8Array
}

export interface SyntheticListDefinition {
  lsid: number
  /** fSimpleList: one level instead of nine */
  simple?: boolean
  levels: SyntheticListLevel[]
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
export function encodeStsh(
  styles: SyntheticStyle[],
  cbStdBase: 0x0a | 0x12 = 0x0a,
  defaultFonts: [number, number, number] = [0, 0, 0],
): Uint8Array {
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
    ...u16le(defaultFonts[0]), // ftcAsci
    ...u16le(defaultFonts[1]), // ftcFE
    ...u16le(defaultFonts[2]), // ftcOther
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

/** an SttbfFfn font table: cData, cbExtra 0, then one FFN (cbFfnM1 + 39 header bytes + UTF-16 name) per font */
export function encodeSttbfFfn(names: string[]): Uint8Array {
  const bytes: number[] = [...u16le(names.length), 0, 0]
  for (const name of names) {
    const ffn: number[] = new Array(39).fill(0)
    for (const char of name) ffn.push(...u16le(char.charCodeAt(0)))
    ffn.push(0, 0)
    bytes.push(ffn.length, ...ffn)
  }
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
export interface SyntheticNote {
  /** body CP of the reference character (0x02 for auto marks, the mark itself for custom ones) */
  cp: number
  /** raw note text; auto notes usually start with \u0002 */
  text: string
  custom?: boolean
}

/** PlcffndRef / PlcfendRef: n+1 CPs plus n FRDs (nAuto; 0 = custom mark) */
function encodeNoteRefs(notes: SyntheticNote[], ccpText: number): Uint8Array {
  const out = new Uint8Array(4 * (notes.length + 1) + 2 * notes.length)
  const view = new DataView(out.buffer)
  notes.forEach((note, index) => view.setUint32(index * 4, note.cp, true))
  view.setUint32(notes.length * 4, ccpText, true)
  notes.forEach((note, index) =>
    view.setUint16(4 * (notes.length + 1) + index * 2, note.custom ? 0 : index + 1, true),
  )
  return out
}

/** the note document text (each note + its paragraph mark, then the guard mark) and its PlcffndTxt CPs */
function encodeNoteStory(notes: SyntheticNote[]): { text: string; cps: number[] } {
  let text = ''
  const cps = [0]
  for (const note of notes) {
    text += note.text + '\r'
    cps.push(text.length)
  }
  text += '\r'
  cps.push(text.length + 2) // Word writes the last CP past the guard mark
  return { text, cps }
}

/** PlcftxbxTxt: CPs + 22-byte FTXBXS entries; the last entry is Word's lid-0 terminator */
export function encodePlcfTxbxTxt(
  stories: Array<{ end: number; spid: number; reusable?: boolean }>,
): Uint8Array {
  const count = stories.length
  const out = new Uint8Array((count + 1) * 4 + count * 22)
  const view = new DataView(out.buffer)
  stories.forEach((story, index) => {
    view.setUint32((index + 1) * 4, story.end, true)
    const base = (count + 1) * 4 + index * 22
    view.setInt32(base, story.spid === 0 ? -1 : 1, true)
    view.setUint16(base + 8, story.reusable ? 1 : 0, true)
    view.setUint32(base + 14, story.spid, true)
  })
  return out
}

export function encodePlcfHdd(cps: number[]): Uint8Array {
  const out = new Uint8Array(4 * cps.length)
  const view = new DataView(out.buffer)
  cps.forEach((cp, index) => view.setUint32(index * 4, cp, true))
  return out
}

/**
 * A synthetic table row (stage 3c-2): grid widths plus the TC descriptor
 * fields of every physical cell. `encodeTableRowPapx` turns it into the TAP
 * sprms a row-end PAPX carries.
 */
export interface SyntheticTableRow {
  /** cell boundaries in twips (deltas become the grid widths) */
  gridWidths: number[]
  cells: Array<{
    /** TC rgf: 0x0001 fFirstMerged, 0x0002 fMerged, 0x0020 fVertMerge, 0x0040 fVertRestart */
    rgf?: number
    /** TC wWidth in twips (informational; the grid widths win) */
    wWidth?: number
    /** visible cell borders (default: all four visible) */
    borders?: { top?: boolean; left?: boolean; bottom?: boolean; right?: boolean }
  }>
  /** sprmTTableHeader: repeat this row on every page */
  header?: boolean
  /** sprmTFCantSplit: keep the row together */
  cantSplit?: boolean
  /** sprmTDyaRowHeight in twips; negative means an exact height */
  rowHeightTwips?: number
  /** sprmTJc: table alignment (only the first row's value is used) */
  alignment?: 'left' | 'center' | 'right'
  /** sprmTTableBorders: emit a visible single-line border set */
  tableBorders?: boolean
  /** sprmTCellPaddingDefault: dxa cell padding in twips (left/right) */
  cellPadding?: number
  /** cell text for the synthetic document builder (cells with fMerged usually stay empty) */
  texts?: string[]
}

/** Encode one sprm; spra 6 uses a u8 length, except TDefTable/sprmPChgTabs (u16, declared = payload + 1). */
export function encodeSprm(opcode: number, operand: Uint8Array | number[]): Uint8Array {
  const bytes = operand instanceof Uint8Array ? operand : Uint8Array.from(operand)
  const spra = (opcode >> 13) & 7
  const header: number[] = [opcode & 0xff, (opcode >> 8) & 0xff]
  if (spra === 6) {
    if (opcode === 0xd608 || opcode === 0xc615) {
      const declared = bytes.length + 1
      header.push(declared & 0xff, (declared >> 8) & 0xff)
    } else {
      if (bytes.length > 255) throw new Error('test fixture: variable sprm operand is too long')
      header.push(bytes.length)
    }
  }
  const out = new Uint8Array(header.length + bytes.length)
  out.set(header, 0)
  out.set(bytes, header.length)
  return out
}

/** One TAP sprm (sprm type 5) with the given operation (0x6000 base + variable length). */
export function encodeTapSprm(operation: number, operand: Uint8Array | number[]): Uint8Array {
  if (operation === 0x08) return encodeSprm(0xd608, operand)
  return encodeSprm(0xd600 | (operation & 0x1ff), operand)
}

/** The TDefTable operand: itcMac, rgdxaCenter, then one 20-byte TC per cell. */
export function encodeTDefTableOperand(row: SyntheticTableRow): Uint8Array {
  const itcMac = row.gridWidths.length
  if (itcMac < 1) throw new Error('test fixture: a table row needs at least one cell')
  const out = new Uint8Array(1 + (itcMac + 1) * 2 + itcMac * 20)
  const view = new DataView(out.buffer)
  out[0] = itcMac
  let cursor = 0
  view.setInt16(1, cursor, true)
  for (let index = 0; index < itcMac; index++) {
    cursor += row.gridWidths[index]!
    view.setInt16(1 + (index + 1) * 2, cursor, true)
  }
  const tcStart = 1 + (itcMac + 1) * 2
  for (let index = 0; index < itcMac; index++) {
    const cell = row.cells[index] ?? {}
    const base = tcStart + index * 20
    view.setUint16(base, cell.rgf ?? 0, true)
    view.setUint16(base + 2, cell.wWidth ?? 0, true)
    const visible = cell.borders ?? { top: true, left: true, bottom: true, right: true }
    const sides = [visible.top, visible.left, visible.bottom, visible.right]
    sides.forEach((on, side) => {
      const offset = base + 4 + side * 4
      out[offset] = on ? 4 : 0 // dptLineWidth
      out[offset + 1] = on ? 1 : 0 // brcType: single / none
      out[offset + 2] = on ? 1 : 0 // ico: auto
      out[offset + 3] = 0 // dptSpace
    })
  }
  return out
}

/** The TAP sprms of one row-end PAPX (fTtp + TDefTable + row properties). */
export function encodeTableRowPapx(row: SyntheticTableRow, itap = 1): Uint8Array {
  const parts: Uint8Array[] = [
    encodeSprm(0x2416, [1]), // sprmPFInTable
    encodeSprm(0x2417, [1]), // sprmPFTtp (row end)
    encodeSprm(0x6649, [itap & 0xff, (itap >> 8) & 0xff, 0, 0]), // sprmPItap
    encodeTapSprm(0x08, encodeTDefTableOperand(row)),
  ]
  if (row.alignment !== undefined) {
    const value = row.alignment === 'center' ? 1 : row.alignment === 'right' ? 2 : 0
    parts.push(encodeTapSprm(0x00, [value, 0]))
  }
  if (row.cantSplit) parts.push(encodeTapSprm(0x03, [1]))
  if (row.header) parts.push(encodeTapSprm(0x04, [1]))
  if (row.rowHeightTwips !== undefined) {
    const value = row.rowHeightTwips
    parts.push(
      encodeTapSprm(0x07, [
        value & 0xff,
        (value >> 8) & 0xff,
        (value >> 16) & 0xff,
        (value >> 24) & 0xff,
      ]),
    )
  }
  if (row.tableBorders) {
    const visible = [4, 1, 1, 0]
    const borders = new Uint8Array(24)
    for (let side = 0; side < 6; side++) borders.set(visible, side * 4)
    parts.push(encodeTapSprm(0x05, borders))
  }
  if (row.cellPadding !== undefined) {
    const width = row.cellPadding
    parts.push(encodeTapSprm(0x34, [0, 0xff, 0x03, 3, width & 0xff, (width >> 8) & 0xff]))
    parts.push(encodeTapSprm(0x34, [0, 0xff, 0x0c, 3, width & 0xff, (width >> 8) & 0xff]))
  }
  const total = parts.reduce((sum, part) => sum + part.length, 0)
  const out = new Uint8Array(total)
  let offset = 0
  for (const part of parts) {
    out.set(part, offset)
    offset += part.length
  }
  return out
}

/** The sprms of a normal table-cell paragraph PAPX. */
export function encodeTableCellPapx(itap = 1): Uint8Array {
  const parts = [
    encodeSprm(0x2416, [1]), // sprmPFInTable
    encodeSprm(0x6649, [itap & 0xff, (itap >> 8) & 0xff, 0, 0]), // sprmPItap
  ]
  const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
  let offset = 0
  for (const part of parts) {
    out.set(part, offset)
    offset += part.length
  }
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

/** Options for a synthetic PICF header + picture blob at a Data stream offset. */
export interface SyntheticPicfOptions {
  /** bytes after the 68-byte PICF header (PNG/JPEG/unsupported payload) */
  payload: Uint8Array
  xExt?: number
  yExt?: number
  dxaGoal?: number
  dyaGoal?: number
  /** scaling factor stored as unsigned 16-bit (48000 = 0.5 inch at 15 twips) */
  mx?: number
  my?: number
  /** override the PICF lcb for corrupt-PICF tests */
  lcb?: number
  /** override the PICF cbHeader for corrupt-PICF tests */
  cbHeader?: number
}

/** A PICF (0x44-byte header) followed by the picture payload. */
export function encodePicf(options: SyntheticPicfOptions): Uint8Array {
  const header = new Uint8Array(0x44)
  const view = new DataView(header.buffer)
  const out = new Uint8Array(header.length + options.payload.length)
  out.set(options.payload, header.length)
  const lcb = options.lcb ?? out.length
  const cbHeader = options.cbHeader ?? 0x44
  view.setUint32(0, lcb >>> 0, true)
  view.setUint16(4, cbHeader, true)
  view.setUint16(6, 0x64, true) // mm constant written by Word/LibreOffice
  view.setInt16(0x08, options.xExt ?? 0, true)
  view.setInt16(0x0a, options.yExt ?? 0, true)
  view.setInt16(0x1c, options.dxaGoal ?? 0, true)
  view.setInt16(0x1e, options.dyaGoal ?? 0, true)
  view.setUint16(0x20, (options.mx ?? 0) & 0xffff, true)
  view.setUint16(0x22, (options.my ?? 0) & 0xffff, true)
  out.set(header, 0)
  return out
}

/** The CHPX sprms of a picture character: sprmCFSpec + sprmCPicLocation. */
export function encodePictureChpx(fcPic: number): Uint8Array {
  const location = new Uint8Array(4)
  new DataView(location.buffer).setInt32(0, fcPic, true)
  const parts = [encodeSprm(0x0855, [1]), encodeSprm(0x6a03, location)]
  const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
  let offset = 0
  for (const part of parts) {
    out.set(part, offset)
    offset += part.length
  }
  return out
}
/** little-endian i16 bytes */
function i16Bytes(value: number): number[] {
  return [value & 0xff, (value >> 8) & 0xff]
}

/** one LVL: LVLF (28 bytes) + grpprlPapx + grpprlChpx + Xst */
export function encodeListLevel(level: SyntheticListLevel): Uint8Array {
  const paragraphSprms =
    level.paragraphSprms ??
    (() => {
      const parts: Uint8Array[] = []
      if (level.indentLeft !== undefined) parts.push(encodeSprm(0x845e, i16Bytes(level.indentLeft)))
      if (level.indentFirstLine !== undefined)
        parts.push(encodeSprm(0x8460, i16Bytes(level.indentFirstLine)))
      if (level.indentRight !== undefined)
        parts.push(encodeSprm(0x845d, i16Bytes(level.indentRight)))
      const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
      let offset = 0
      for (const part of parts) {
        out.set(part, offset)
        offset += part.length
      }
      return out
    })()
  const characterSprms = level.characterSprms ?? new Uint8Array(0)
  const textCodes: number[] = []
  for (const char of level.levelText ?? '') textCodes.push(char.charCodeAt(0))
  const lvlf = new Uint8Array(28)
  const view = new DataView(lvlf.buffer)
  view.setUint32(0, level.startAt ?? 1, true)
  lvlf[4] = level.numberFormat ?? 0
  lvlf[5] =
    ((level.alignment ?? 0) & 0x03) | (level.legal ? 0x04 : 0) | (level.noRestart ? 0x08 : 0)
  lvlf[0x0f] = level.follow ?? 0
  lvlf[0x18] = characterSprms.length
  lvlf[0x19] = paragraphSprms.length
  const out = new Uint8Array(
    28 + paragraphSprms.length + characterSprms.length + 2 + textCodes.length * 2,
  )
  out.set(lvlf, 0)
  out.set(paragraphSprms, 28)
  out.set(characterSprms, 28 + paragraphSprms.length)
  const xstStart = 28 + paragraphSprms.length + characterSprms.length
  new DataView(out.buffer).setUint16(xstStart, textCodes.length, true)
  textCodes.forEach((code, index) =>
    new DataView(out.buffer).setUint16(xstStart + 2 + index * 2, code, true),
  )
  return out
}

/**
 * A PlfLst split into its header (cLst + LSTF records, what lcbPlfLst covers)
 * and the level array that follows it in the table stream.
 */
export function encodePlfLst(lists: SyntheticListDefinition[]): {
  header: Uint8Array
  levels: Uint8Array
} {
  const header = new Uint8Array(2 + lists.length * 28)
  const headerView = new DataView(header.buffer)
  headerView.setUint16(0, lists.length, true)
  lists.forEach((list, index) => {
    const base = 2 + index * 28
    headerView.setUint32(base, list.lsid, true)
    header[base + 0x1a] = list.simple ? 0x01 : 0x00
  })
  const encoded = lists.map((list) => {
    const levels = [...list.levels]
    if (!list.simple) {
      // a non-simple list always carries nine LVLs; pad the sparse test input
      while (levels.length < 9)
        levels.push({ numberFormat: 0, levelText: String.fromCharCode(levels.length) + '.' })
    }
    return levels.map(encodeListLevel)
  })
  const levels = new Uint8Array(
    encoded.reduce((sum, list) => sum + list.reduce((inner, level) => inner + level.length, 0), 0),
  )
  let offset = 0
  for (const list of encoded)
    for (const level of list) {
      levels.set(level, offset)
      offset += level.length
    }
  return { header, levels }
}

/** a PlfLfo: one LFO per lsid (clfolvl 0) plus the LFOData cp records */
/** One LFOLVL of a synthetic list override. */
export interface SyntheticLevelOverride {
  ilvl: number
  startAt?: number
  level?: SyntheticListLevel
}

/** PlfLfo: lfoMac, the LFO array, then one LFOData (cp + LFOLVLs) per LFO. */
export function encodePlfLfo(
  lsids: number[],
  levelOverrides: Array<SyntheticLevelOverride[] | undefined> = [],
): Uint8Array {
  const header = new Uint8Array(4 + lsids.length * 16)
  const headerView = new DataView(header.buffer)
  headerView.setUint32(0, lsids.length, true)
  lsids.forEach((lsid, index) => {
    headerView.setUint32(4 + index * 16, lsid, true)
    header[4 + index * 16 + 0x0c] = levelOverrides[index]?.length ?? 0
  })
  const data: Uint8Array[] = []
  lsids.forEach((_, index) => {
    const cp = new Uint8Array(4)
    new DataView(cp.buffer).setUint32(0, 0xffffffff, true)
    data.push(cp)
    for (const override of levelOverrides[index] ?? []) {
      const lfolvl = new Uint8Array(8)
      new DataView(lfolvl.buffer).setUint32(0, override.startAt ?? 0, true)
      lfolvl[4] =
        (override.ilvl & 0x0f) |
        (override.startAt !== undefined ? 0x10 : 0) |
        (override.level ? 0x20 : 0)
      data.push(lfolvl)
      if (override.level) data.push(encodeListLevel(override.level))
    }
  })
  const out = new Uint8Array(header.length + data.reduce((sum, part) => sum + part.length, 0))
  out.set(header, 0)
  let offset = header.length
  for (const part of data) {
    out.set(part, offset)
    offset += part.length
  }
  return out
}

/** the PAPX sprms of a list paragraph: sprmPIlvl + sprmPIlfo (ilfo is 1-based) */
export function encodeListParagraphPapx(ilfo: number, ilvl = 0): Uint8Array {
  const parts = [encodeSprm(0x260a, [ilvl]), encodeSprm(0x460b, [ilfo & 0xff, (ilfo >> 8) & 0xff])]
  const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
  let offset = 0
  for (const part of parts) {
    out.set(part, offset)
    offset += part.length
  }
  return out
}
/** wrap streams in an OLE2 compound file (cfb seeds its own producer marker, ignored on read) */
export function buildCompound(streams: Array<[string, Uint8Array]>): Uint8Array {
  const container = CFB.utils.cfb_new()
  for (const [name, content] of streams) CFB.utils.cfb_add(container, name, Array.from(content))
  const written = CFB.write(container, { type: 'buffer', fileType: 'cfb' }) as Uint8Array
  return new Uint8Array(written)
}

function concatBytes(parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
  let offset = 0
  for (const part of parts) {
    out.set(part, offset)
    offset += part.length
  }
  return out
}

function officeArtRecord(
  type: number,
  instance: number,
  payload: Uint8Array,
  container = false,
  lengthOverride?: number,
): Uint8Array {
  const out = new Uint8Array(8 + payload.length)
  const view = new DataView(out.buffer)
  view.setUint16(0, (instance << 4) | (container ? 0x0f : 0), true)
  view.setUint16(2, type, true)
  view.setUint32(4, lengthOverride ?? payload.length, true)
  out.set(payload, 8)
  return out
}

const BLIP_RECORDS: Record<
  SyntheticBlipKind,
  { type: number; instance: number; metafile: boolean }
> = {
  emf: { type: 0xf01a, instance: 0x3d4, metafile: true },
  wmf: { type: 0xf01b, instance: 0x216, metafile: true },
  pict: { type: 0xf01c, instance: 0x542, metafile: true },
  jpeg: { type: 0xf01d, instance: 0x46a, metafile: false },
  png: { type: 0xf01e, instance: 0x6e0, metafile: false },
  dib: { type: 0xf01f, instance: 0x7a8, metafile: false },
}

/** one OfficeArtBlip record: UID, then tag + bytes (bitmaps) or the 34-byte metafile header + data */
export function encodeBlipRecord(blip: SyntheticBlip): Uint8Array {
  const kind = BLIP_RECORDS[blip.kind]
  const uid = new Uint8Array(16).fill(0xab)
  let payload: Uint8Array
  if (kind.metafile) {
    const data = blip.uncompressed ? blip.bytes : new Uint8Array(deflateSync(blip.bytes))
    const header = new Uint8Array(34)
    const view = new DataView(header.buffer)
    view.setUint32(0, blip.bytes.length, true)
    view.setUint32(28, data.length, true)
    header[32] = blip.uncompressed ? 0xfe : 0
    header[33] = 0xfe
    payload = concatBytes([uid, header, data])
  } else payload = concatBytes([uid, Uint8Array.from([0xff]), blip.bytes])
  return officeArtRecord(kind.type, kind.instance, payload, false, blip.recordLength)
}

/** FBSE: btWin32 / btMac, UID, tag, size, cRef, foDelay, name length, then an embedded blip when asked */
function encodeFbse(blip: SyntheticBlip, foDelay: number, embedded?: Uint8Array): Uint8Array {
  const body = new Uint8Array(36)
  const view = new DataView(body.buffer)
  const bt: Record<SyntheticBlipKind, number> = { emf: 2, wmf: 3, pict: 4, jpeg: 5, png: 6, dib: 7 }
  body[0] = bt[blip.kind]
  body[1] = bt[blip.kind]
  body.fill(0xab, 2, 18)
  view.setUint16(18, 0xff, true)
  view.setUint32(20, (embedded?.length ?? 0) || blip.bytes.length, true)
  view.setUint32(24, 1, true)
  view.setUint32(28, foDelay, true)
  const payload = embedded ? concatBytes([body, embedded]) : body
  return officeArtRecord(0xf007, bt[blip.kind], payload)
}

function rectBytes(rect: [number, number, number, number]): Uint8Array {
  const out = new Uint8Array(16)
  const view = new DataView(out.buffer)
  rect.forEach((value, index) => view.setInt32(index * 4, value, true))
  return out
}

function optRecord(entries: Array<[number, number]>): Uint8Array {
  const out = new Uint8Array(entries.length * 6)
  const view = new DataView(out.buffer)
  entries.forEach(([pid, value], index) => {
    view.setUint16(index * 6, pid, true)
    view.setUint32(index * 6 + 2, value, true)
  })
  return officeArtRecord(0xf00b, entries.length, out)
}

function fspRecord(spid: number, flags: number, shapeType: number): Uint8Array {
  const out = new Uint8Array(8)
  const view = new DataView(out.buffer)
  view.setUint32(0, spid, true)
  view.setUint32(4, flags, true)
  return officeArtRecord(0xf00a, shapeType, out)
}

/**
 * The DggInfo bytes plus the blip data that goes into the WordDocument stream
 * (`delayBlips`, to be placed at `delayBase`), and the PlcSpa tables.
 */
export function encodeFloatingShapes(
  shapes: SyntheticFloatingShape[],
  delayBase: number,
): { dggInfo: Uint8Array; delayBlips: Uint8Array; plcSpaMom?: Uint8Array; plcSpaHdr?: Uint8Array } {
  const fbses: Uint8Array[] = []
  const delayParts: Uint8Array[] = []
  let delayOffset = delayBase
  const registerBlip = (blip: SyntheticBlip): number => {
    const record = encodeBlipRecord(blip)
    if (blip.embedded) fbses.push(encodeFbse(blip, 0xffffffff, record))
    else {
      fbses.push(encodeFbse(blip, delayOffset))
      delayParts.push(record)
      delayOffset += record.length
    }
    return fbses.length
  }
  let nextSpid = 1025
  const spContainer = (
    spid: number,
    flags: number,
    blip: SyntheticBlip | undefined,
    pib: number | undefined,
    extra: Uint8Array[],
    props: Array<[number, number]> = [],
  ): Uint8Array => {
    const pibValue = pib ?? (blip ? registerBlip(blip) : undefined)
    const parts = [
      fspRecord(spid, flags, pibValue !== undefined ? 75 : flags & 1 ? 0 : 202),
      ...(pibValue !== undefined
        ? [optRecord([[0x4104, pibValue], ...props])]
        : [optRecord([[0x01bf, 0x00010000], ...props])]),
      ...extra,
    ]
    return officeArtRecord(0xf004, 0, concatBytes(parts), true)
  }
  const groupContainer = (
    spid: number,
    flags: number,
    group: { bounds: [number, number, number, number]; children: SyntheticShapeChild[] },
    childAnchor?: [number, number, number, number],
  ): Uint8Array => {
    const self = spContainer(spid, flags | 0x0001, undefined, undefined, [
      officeArtRecord(0xf009, 1, rectBytes(group.bounds)),
      ...(childAnchor ? [officeArtRecord(0xf00f, 0, rectBytes(childAnchor))] : []),
    ])
    const members = group.children.map((child) => {
      const childSpid = nextSpid++
      if (child.group) return groupContainer(childSpid, 0x0002, child.group, child.anchor)
      return spContainer(childSpid, 0x0a02, child.blip, undefined, [
        officeArtRecord(0xf00f, 0, rectBytes(child.anchor)),
      ])
    })
    return officeArtRecord(0xf003, 0, concatBytes([self, ...members]), true)
  }
  const drawings = new Map<number, Uint8Array[]>()
  const anchors = new Map<number, Array<{ shape: SyntheticFloatingShape; spid: number }>>()
  for (const shape of shapes) {
    const label = shape.story === 'header' ? 1 : 0
    const spid = shape.spid ?? nextSpid++
    const container = shape.group
      ? groupContainer(spid, 0x0a00, shape.group)
      : spContainer(
          spid,
          0x0a00,
          shape.blip,
          shape.pib,
          [officeArtRecord(0xf010, 0, new Uint8Array(4))],
          shape.props,
        )
    drawings.set(label, [...(drawings.get(label) ?? []), container])
    anchors.set(label, [...(anchors.get(label) ?? []), { shape, spid }])
  }
  const fdgg = new Uint8Array(16 + 8)
  new DataView(fdgg.buffer).setUint32(0, nextSpid, true)
  const bstore = officeArtRecord(0xf001, fbses.length, concatBytes(fbses), true)
  const dgg = officeArtRecord(
    0xf000,
    0,
    concatBytes([officeArtRecord(0xf006, 0, fdgg), bstore]),
    true,
  )
  const wordDrawings: Uint8Array[] = []
  for (const label of [0, 1]) {
    const containers = drawings.get(label)
    if (!containers) continue
    const fdg = new Uint8Array(8)
    new DataView(fdg.buffer).setUint32(0, containers.length + 1, true)
    const patriarch = officeArtRecord(
      0xf004,
      0,
      concatBytes([
        officeArtRecord(0xf009, 1, new Uint8Array(16)),
        fspRecord(1024 + label, 0x0005, 0),
      ]),
      true,
    )
    const spgr = officeArtRecord(0xf003, 0, concatBytes([patriarch, ...containers]), true)
    wordDrawings.push(
      Uint8Array.from([label]),
      officeArtRecord(
        0xf002,
        label + 1,
        concatBytes([officeArtRecord(0xf008, 0, fdg), spgr]),
        true,
      ),
    )
  }
  const plcSpa = (label: number): Uint8Array | undefined => {
    const entries = anchors.get(label)
    if (!entries) return undefined
    const out = new Uint8Array((entries.length + 1) * 4 + entries.length * 26)
    const view = new DataView(out.buffer)
    entries.forEach(({ shape }, index) => view.setUint32(index * 4, shape.cp, true))
    view.setUint32(entries.length * 4, (entries.at(-1)?.shape.cp ?? 0) + 1, true)
    entries.forEach(({ shape, spid }, index) => {
      const base = (entries.length + 1) * 4 + index * 26
      view.setUint32(base, spid, true)
      shape.rect.forEach((value, at) => view.setInt32(base + 4 + at * 4, value, true))
      const flags =
        ((shape.bx ?? 0) << 1) |
        ((shape.by ?? 0) << 3) |
        ((shape.wr ?? 2) << 5) |
        (shape.belowText ? 0x4000 : 0)
      view.setUint16(base + 20, flags, true)
    })
    return out
  }
  return {
    dggInfo: concatBytes([dgg, ...wordDrawings]),
    delayBlips: concatBytes(delayParts),
    ...(plcSpa(0) ? { plcSpaMom: plcSpa(0) } : {}),
    ...(plcSpa(1) ? { plcSpaHdr: plcSpa(1) } : {}),
  }
}

/** an extended STTB (0xFFFF, UTF-16 strings, cbExtra bytes after each string) */
export function encodeSttb(strings: string[], extra: Uint8Array[] = [], cbExtra = 0): Uint8Array {
  const parts: number[] = [0xff, 0xff, ...u16le(strings.length), ...u16le(cbExtra)]
  strings.forEach((text, index) => {
    parts.push(...u16le(text.length))
    for (const char of text) parts.push(...u16le(char.charCodeAt(0)))
    parts.push(...(extra[index] ?? new Uint8Array(cbExtra)))
  })
  return Uint8Array.from(parts)
}

interface CommentTables {
  plcfandRef: Uint8Array
  plcfandTxt: Uint8Array
  sttbfRMark: Uint8Array
  sttbfAtnBkmk?: Uint8Array
  plcfAtnBkf?: Uint8Array
  plcfAtnBkl?: Uint8Array
  atrdExtra?: Uint8Array
}

/** the comment-document text plus every table PlcfandRef-based comment reading needs */
function encodeComments(comments: SyntheticComment[]): { text: string; tables: CommentTables } {
  const authors = [...new Set(comments.map((comment) => comment.author))]
  let text = ''
  const txtCps = [0]
  for (const comment of comments) {
    const paragraphs = Array.isArray(comment.text) ? comment.text : [comment.text]
    text += '\u0005' + paragraphs.join('\r') + '\r'
    txtCps.push(text.length)
  }
  txtCps.push(text.length + 1)
  const u32 = (value: number): number[] => [
    value & 0xff,
    (value >>> 8) & 0xff,
    (value >>> 16) & 0xff,
    (value >>> 24) & 0xff,
  ]
  const ranged = comments
    .map((comment, index) => ({ comment, index }))
    .filter((entry) => entry.comment.range !== undefined)
  const tagOf = (index: number): number => 0x1000 + index
  const ref: number[] = []
  for (const comment of comments) ref.push(...u32(comment.referenceCp))
  ref.push(...u32((comments[comments.length - 1]?.referenceCp ?? 0) + 1))
  comments.forEach((comment, index) => {
    const xst = new Array<number>(20).fill(0)
    xst.splice(0, 2, ...u16le(comment.initials.length))
    ;[...comment.initials].slice(0, 9).forEach((char, position) => {
      xst.splice(2 + position * 2, 2, ...u16le(char.charCodeAt(0)))
    })
    ref.push(
      ...xst,
      ...u16le(authors.indexOf(comment.author)),
      ...u16le(0),
      ...u16le(0),
      ...u32(comment.range ? tagOf(index) : -1),
    )
  })
  const tables: CommentTables = {
    plcfandRef: Uint8Array.from(ref),
    plcfandTxt: Uint8Array.from(txtCps.flatMap(u32)),
    sttbfRMark: encodeSttb(authors),
  }
  if (ranged.length > 0) {
    // bookmark starts sorted by CP select the BKF order; ends sorted by CP select the BKL order
    const byStart = [...ranged].sort(
      (left, right) => left.comment.range![0] - right.comment.range![0],
    )
    const byEnd = [...ranged].sort(
      (left, right) => left.comment.range![1] - right.comment.range![1],
    )
    const lastCp = Math.max(...ranged.map((entry) => entry.comment.range![1])) + 1
    const bkf: number[] = []
    for (const entry of byStart) bkf.push(...u32(entry.comment.range![0]))
    bkf.push(...u32(lastCp))
    for (const entry of byStart) bkf.push(...u16le(byEnd.indexOf(entry)), ...u16le(0))
    const bkl: number[] = []
    for (const entry of byEnd) bkl.push(...u32(entry.comment.range![1]))
    bkl.push(...u32(lastCp))
    tables.plcfAtnBkf = Uint8Array.from(bkf)
    tables.plcfAtnBkl = Uint8Array.from(bkl)
    tables.sttbfAtnBkmk = encodeSttb(
      byStart.map(() => ''),
      byStart.map((entry) =>
        Uint8Array.from([...u16le(0x100), ...u32(tagOf(entry.index)), ...u32(-1)]),
      ),
      10,
    )
  }
  if (comments.some((comment) => comment.dttm !== undefined)) {
    const extra: number[] = []
    for (const comment of comments) extra.push(...u32(comment.dttm ?? 0), ...new Array(14).fill(0))
    tables.atrdExtra = Uint8Array.from(extra)
  }
  return { text, tables }
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
  plcffndRef?: { fc: number; lcb: number }
  plcffndTxt?: { fc: number; lcb: number }
  plcfendRef?: { fc: number; lcb: number }
  plcfendTxt?: { fc: number; lcb: number }
  plcflst?: { fc: number; lcb: number }
  plcflfo?: { fc: number; lcb: number }
  sttbfFfn?: { fc: number; lcb: number }
  plcSpaMom?: { fc: number; lcb: number }
  plcSpaHdr?: { fc: number; lcb: number }
  dggInfo?: { fc: number; lcb: number }
  /** comment tables by FIB pair index */
  extraPairs?: Array<[number, { fc: number; lcb: number }]>
  plcfTxbxTxt?: { fc: number; lcb: number }
  plcfHdrTxbxTxt?: { fc: number; lcb: number }
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
  const PLCFLST_PAIR_INDEX = 73
  const PLCFLFO_PAIR_INDEX = 74
  const STTBF_FFN_PAIR_INDEX = 15
  if (values.stshf) {
    view.setUint32(offset + STSHF_PAIR_INDEX * 8, values.stshf.fc, true)
    view.setUint32(offset + STSHF_PAIR_INDEX * 8 + 4, values.stshf.lcb, true)
  }
  for (const [index, pair] of [
    [PLCFSED_PAIR_INDEX, values.plcfsed],
    [PLCFHDD_PAIR_INDEX, values.plcfhdd],
    [2, values.plcffndRef],
    [3, values.plcffndTxt],
    [46, values.plcfendRef],
    [47, values.plcfendTxt],
    [PLCFLST_PAIR_INDEX, values.plcflst],
    [PLCFLFO_PAIR_INDEX, values.plcflfo],
    [STTBF_FFN_PAIR_INDEX, values.sttbfFfn],
    [40, values.plcSpaMom],
    [41, values.plcSpaHdr],
    [50, values.dggInfo],
    [56, values.plcfTxbxTxt],
    [58, values.plcfHdrTxbxTxt],
    ...(values.extraPairs ?? []),
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
  let plfLst = options.plfLst
  let plfLstLevels: Uint8Array = new Uint8Array(0)
  let shapeTables: Uint8Array = new Uint8Array(0)
  let noteTableBytes: Uint8Array = new Uint8Array(0)
  if (!plfLst && options.lists && options.lists.length > 0) {
    const encoded = encodePlfLst(options.lists)
    plfLst = encoded.header
    plfLstLevels = encoded.levels
  }
  const plfLfo =
    options.plfLfo ??
    (options.lists && options.lists.length > 0
      ? encodePlfLfo(
          options.listOverrides ?? options.lists.map((list) => list.lsid),
          options.listLevelOverrides ?? [],
        )
      : undefined)
  const stsh =
    options.stsh ??
    (options.stshStyles
      ? encodeStsh(options.stshStyles, options.stshCbStdBase, options.stshDefaultFonts)
      : undefined)
  const sttbfFfn = options.sttbfFfn ?? (options.fonts ? encodeSttbfFfn(options.fonts) : undefined)
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
    const footnoteStory = options.footnotes ? encodeNoteStory(options.footnotes) : undefined
    const endnoteStory = options.endnotes ? encodeNoteStory(options.endnotes) : undefined
    const addStory = (text: string): void => {
      pcds.push({ fc: byteCursor, compressed: false })
      const bytes = encodeUtf16Le(text)
      encoded.push(bytes)
      byteCursor += bytes.length
      charCount += text.length
      cps.push(charCount)
    }
    if (footnoteStory) addStory(footnoteStory.text)
    if (sectionDefs.length > 0) {
      if ((options.ccpFtn ?? 0) > 0 && !footnoteStory)
        throw new Error(
          'test fixture: a ccpFtn count without footnote text cannot be combined with headers',
        )
      hddCps.push(0, 0, 0, 0, 0, 0, 0) // the six empty footnote/endnote separator stories
      for (const stories of storySets) {
        for (let variant = 0; variant < 6; variant++) {
          const text = stories[variant]
          // '' is an unlinked blank story (just its paragraph mark); undefined is linked to previous
          if (text !== undefined) {
            headerText += text + '\r'
            headerChars += text.length + 1
          }
          hddCps.push(headerChars)
        }
      }
      hddCps.push(headerChars + 1) // n+2nd CP (one past the last story limit)
    }
    const comments = options.comments ? encodeComments(options.comments) : undefined
    if (headerText !== '') addStory(headerText)
    if (comments) addStory(comments.text)
    if (endnoteStory) addStory(endnoteStory.text)
    // text box documents: body stories, then header stories, each closed by a terminator mark
    const textBoxDocument = (story: 'body' | 'header'): { text: string; table?: Uint8Array } => {
      const stories = (options.textBoxStories ?? []).filter(
        (entry) => (entry.story ?? 'body') === story,
      )
      if (stories.length === 0) return { text: '' }
      let text = ''
      const entries: Array<{ end: number; spid: number; reusable?: boolean }> = []
      for (const entry of stories) {
        text += entry.text + '\r'
        entries.push({ end: text.length, spid: entry.spid, reusable: entry.reusable })
      }
      text += '\r'
      entries.push({ end: text.length, spid: 0 })
      return { text, table: encodePlcfTxbxTxt(entries) }
    }
    const bodyTextBoxes = textBoxDocument('body')
    const headerTextBoxes = textBoxDocument('header')
    const textBoxText = bodyTextBoxes.text + headerTextBoxes.text
    if (textBoxText !== '') addStory(textBoxText)
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
    const shapes = options.floatingShapes
      ? encodeFloatingShapes(options.floatingShapes, sepxStart + sepxTotal)
      : undefined
    plcfsed ??= sectionDefs.length > 0 ? encodeSectionPlcf(sectionCps, sepxFcs) : undefined
    plcfhdd ??=
      sectionDefs.length > 0 && (storiesProvided || headerText !== '')
        ? encodePlcfHdd(hddCps)
        : undefined
    const noteTables = {
      ftnRef: options.footnotes ? encodeNoteRefs(options.footnotes, bodyCharCount) : undefined,
      ftnTxt:
        footnoteStory && !options.omitNoteTextTables ? encodePlcfHdd(footnoteStory.cps) : undefined,
      endRef: options.endnotes ? encodeNoteRefs(options.endnotes, bodyCharCount) : undefined,
      endTxt:
        endnoteStory && !options.omitNoteTextTables ? encodePlcfHdd(endnoteStory.cps) : undefined,
    }
    noteTableBytes = concatBytes(
      [noteTables.ftnRef, noteTables.ftnTxt, noteTables.endRef, noteTables.endTxt].map(
        (bytes) => bytes ?? new Uint8Array(0),
      ),
    )
    // table layout: CLX, formatting tables, stylesheet, section and header tables
    const baseTableLength = fcClx + clx.length
    const chpxPlcFc = baseTableLength
    const papxPlcFc = chpxPlcFc + (chpxPlcf?.length ?? 0)
    const stshFc = papxPlcFc + (papxPlcf?.length ?? 0)
    const plcfsedFc = stshFc + (stsh?.length ?? 0)
    const plcfhddFc = plcfsedFc + (plcfsed?.length ?? 0)
    const ftnRefFc = plcfhddFc + (plcfhdd?.length ?? 0)
    const ftnTxtFc = ftnRefFc + (noteTables.ftnRef?.length ?? 0)
    const endRefFc = ftnTxtFc + (noteTables.ftnTxt?.length ?? 0)
    const endTxtFc = endRefFc + (noteTables.endRef?.length ?? 0)
    const plcflstFc = endTxtFc + (noteTables.endTxt?.length ?? 0)
    const plfLstLevelsFc = plcflstFc + (plfLst?.length ?? 0)
    const plcflfoFc = plfLstLevelsFc + plfLstLevels.length
    const sttbfFfnFc = plcflfoFc + (plfLfo?.length ?? 0)
    const dggInfoFc = sttbfFfnFc + (sttbfFfn?.length ?? 0)
    const plcSpaMomFc = dggInfoFc + (shapes?.dggInfo.length ?? 0)
    const plcSpaHdrFc = plcSpaMomFc + (shapes?.plcSpaMom?.length ?? 0)
    const plcfTxbxTxtFc = plcSpaHdrFc + (shapes?.plcSpaHdr?.length ?? 0)
    const plcfHdrTxbxTxtFc = plcfTxbxTxtFc + (bodyTextBoxes.table?.length ?? 0)
    shapeTables = concatBytes([
      shapes?.dggInfo ?? new Uint8Array(0),
      shapes?.plcSpaMom ?? new Uint8Array(0),
      shapes?.plcSpaHdr ?? new Uint8Array(0),
      bodyTextBoxes.table ?? new Uint8Array(0),
      headerTextBoxes.table ?? new Uint8Array(0),
    ])
    // comment tables follow the shape and text box tables at the end of the table stream
    const commentPairs: Array<[number, { fc: number; lcb: number }]> = []
    const commentParts: Uint8Array[] = []
    if (comments) {
      let cursor = plcfHdrTxbxTxtFc + (headerTextBoxes.table?.length ?? 0)
      const pairIndex: Record<keyof CommentTables, number> = {
        plcfandRef: 4,
        plcfandTxt: 5,
        sttbfAtnBkmk: 37,
        plcfAtnBkf: 42,
        plcfAtnBkl: 43,
        sttbfRMark: 51,
        atrdExtra: 112,
      }
      for (const key of Object.keys(pairIndex) as Array<keyof CommentTables>) {
        const bytes = comments.tables[key]
        if (!bytes) continue
        const lcb = key === 'plcfandRef' ? (options.lcbPlcfandRef ?? bytes.length) : bytes.length
        if (!(key === 'plcfandTxt' && options.omitPlcfandTxt))
          commentPairs.push([pairIndex[key], { fc: cursor, lcb }])
        commentParts.push(bytes)
        cursor += bytes.length
      }
      shapeTables = concatBytes([shapeTables, ...commentParts])
    }
    const word = new Uint8Array(
      nextPage * SYNTHETIC_FKP_PAGE_BYTES + sepxTotal + (shapes?.delayBlips.length ?? 0),
    )
    if (shapes) word.set(shapes.delayBlips, sepxStart + sepxTotal)
    writeFib(word, {
      wIdent,
      nFib,
      flags,
      ccpText: options.ccpText ?? bodyCharCount,
      fcClx,
      lcbClx: options.lcbClx ?? clx.length,
      ccpFtn: options.ccpFtn ?? footnoteStory?.text.length,
      ccpHdd: options.ccpHdd ?? headerChars,
      ccpAtn: options.ccpAtn ?? comments?.text.length,
      ccpEdn: options.ccpEdn ?? endnoteStory?.text.length,
      ccpTxbx: options.ccpTxbx ?? (bodyTextBoxes.text ? bodyTextBoxes.text.length : undefined),
      ccpHdrTxbx:
        options.ccpHdrTxbx ?? (headerTextBoxes.text ? headerTextBoxes.text.length : undefined),
      ...(bodyTextBoxes.table
        ? { plcfTxbxTxt: { fc: plcfTxbxTxtFc, lcb: bodyTextBoxes.table.length } }
        : {}),
      ...(headerTextBoxes.table
        ? { plcfHdrTxbxTxt: { fc: plcfHdrTxbxTxtFc, lcb: headerTextBoxes.table.length } }
        : {}),
      extraPairs: commentPairs,
      ...(chpxPlcf
        ? { chpxPlcf: { fc: chpxPlcFc, lcb: options.chpxPlcLcb ?? chpxPlcf.length } }
        : {}),
      ...(papxPlcf
        ? { papxPlcf: { fc: papxPlcFc, lcb: options.papxPlcLcb ?? papxPlcf.length } }
        : {}),
      ...(stsh ? { stshf: { fc: stshFc, lcb: options.lcbStshf ?? stsh.length } } : {}),
      ...(plcfsed ? { plcfsed: { fc: plcfsedFc, lcb: options.lcbPlcfsed ?? plcfsed.length } } : {}),
      ...(plcfhdd ? { plcfhdd: { fc: plcfhddFc, lcb: options.lcbPlcfhdd ?? plcfhdd.length } } : {}),
      ...(noteTables.ftnRef ? { plcffndRef: { fc: ftnRefFc, lcb: noteTables.ftnRef.length } } : {}),
      ...(noteTables.ftnTxt
        ? { plcffndTxt: { fc: ftnTxtFc, lcb: options.lcbPlcffndTxt ?? noteTables.ftnTxt.length } }
        : {}),
      ...(noteTables.endRef ? { plcfendRef: { fc: endRefFc, lcb: noteTables.endRef.length } } : {}),
      ...(noteTables.endTxt ? { plcfendTxt: { fc: endTxtFc, lcb: noteTables.endTxt.length } } : {}),
      ...(plfLst ? { plcflst: { fc: plcflstFc, lcb: options.lcbPlfLst ?? plfLst.length } } : {}),
      ...(plfLfo ? { plcflfo: { fc: plcflfoFc, lcb: options.lcbPlfLfo ?? plfLfo.length } } : {}),
      ...(sttbfFfn
        ? { sttbfFfn: { fc: sttbfFfnFc, lcb: options.lcbSttbfFfn ?? sttbfFfn.length } }
        : {}),
      ...(shapes ? { dggInfo: { fc: dggInfoFc, lcb: shapes.dggInfo.length } } : {}),
      ...(shapes?.plcSpaMom
        ? { plcSpaMom: { fc: plcSpaMomFc, lcb: shapes.plcSpaMom.length } }
        : {}),
      ...(shapes?.plcSpaHdr
        ? { plcSpaHdr: { fc: plcSpaHdrFc, lcb: shapes.plcSpaHdr.length } }
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
        (plcfhdd?.length ?? 0) +
        noteTableBytes.length +
        (plfLst?.length ?? 0) +
        plfLstLevels.length +
        (plfLfo?.length ?? 0) +
        (sttbfFfn?.length ?? 0) +
        shapeTables.length,
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
    if (plcfhdd) {
      table.set(plcfhdd, cursor)
      cursor += plcfhdd.length
    }
    if (noteTableBytes.length > 0) {
      table.set(noteTableBytes, cursor)
      cursor += noteTableBytes.length
    }
    if (plfLst) {
      table.set(plfLst, cursor)
      cursor += plfLst.length
    }
    if (plfLstLevels.length > 0) {
      table.set(plfLstLevels, cursor)
      cursor += plfLstLevels.length
    }
    if (plfLfo) {
      table.set(plfLfo, cursor)
      cursor += plfLfo.length
    }
    if (sttbfFfn) {
      table.set(sttbfFfn, cursor)
      cursor += sttbfFfn.length
    }
    if (shapeTables.length > 0) table.set(shapeTables, cursor)
  }

  const streams: Array<[string, Uint8Array]> = []
  if (!options.omitWordDocument) streams.push(['WordDocument', wordDocument])
  if (!options.omitTableStream) streams.push([tableStream, table])
  if (options.dataStream) streams.push(['Data', options.dataStream])
  return buildCompound(streams)
}
