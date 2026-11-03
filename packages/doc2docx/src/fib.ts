import { LegacyDocError } from './errors'

/** FibBase.wIdent: every Word 97-2003 document starts with this value */
export const FIB_IDENT = 0xa5ec
/** oldest format version this reader supports (Word 97) */
export const MIN_SUPPORTED_NFIB = 0x00c1
/** newest format version this reader supports (Word 2003 and the Word 97 format later versions write) */
export const MAX_SUPPORTED_NFIB = 0x0112

const FIB_BASE_BYTES = 32
/** FibBase.flags bits */
const FLAG_ENCRYPTED = 0x0100
const FLAG_WHICH_TABLE_STREAM = 0x0200
const FLAG_OBFUSCATED = 0x8000

/** one fc/lcb pair of FibRgFcLcb97: a byte range in the table stream */
export interface FibPair {
  fc: number
  lcb: number
}

export interface FibInfo {
  /** format version (0x00C1 = Word 97 .. 0x0112 = Word 2003+) */
  nFib: number
  /** raw FibBase.flags */
  flags: number
  encrypted: boolean
  obfuscated: boolean
  /** the stream that holds the CLX, STSH, PAPX/CHPX and other tables */
  tableStream: '0Table' | '1Table'
  /** characters in the main document story (this stage reads only this range) */
  ccpText: number
  ccpFtn: number
  ccpHdd: number
  ccpAtn: number
  ccpEdn: number
  ccpTxbx: number
  ccpHdrTxbx: number
  /** piece table (CLX) location in the table stream */
  fcClx: number
  lcbClx: number
  /** stylesheet (stage 3) */
  fcStshf: number
  lcbStshf: number
  /** section descriptor table (stage 3c) */
  fcPlcfsed: number
  lcbPlcfsed: number
  /** header/footer story table (stage 3c) */
  fcPlcfhdd: number
  lcbPlcfhdd: number
  /** footnote reference / text tables (PlcffndRef over the body, PlcffndTxt over the footnote document) */
  fcPlcffndRef: number
  lcbPlcffndRef: number
  fcPlcffndTxt: number
  lcbPlcffndTxt: number
  /** endnote reference / text tables (PlcfendRef / PlcfendTxt) */
  fcPlcfendRef: number
  lcbPlcfendRef: number
  fcPlcfendTxt: number
  lcbPlcfendTxt: number
  /** list definitions (stage 3c-4 PlfLst; the level array follows the LSTFs) */
  fcPlfLst: number
  lcbPlfLst: number
  /** list format overrides (stage 3c-4 PlfLfo) */
  fcPlfLfo: number
  lcbPlfLfo: number
  /** font table (SttbfFfn): ftc -> font name */
  fcSttbfFfn: number
  lcbSttbfFfn: number
  /** CHPX / PAPX bin tables (stage 2) */
  fcPlcfbteChpx: number
  lcbPlcfbteChpx: number
  fcPlcfbtePapx: number
  lcbPlcfbtePapx: number
  /** document properties (stage 2) */
  fcDop: number
  lcbDop: number
  /** floating shape anchors of the body (PlcSpaMom) and the header document (PlcSpaHdr) */
  fcPlcSpaMom: number
  lcbPlcSpaMom: number
  fcPlcSpaHdr: number
  lcbPlcSpaHdr: number
  /** OfficeArt drawing data (DggInfo): blip store + shape containers */
  fcDggInfo: number
  lcbDggInfo: number
  /** comments: anchors (PlcfandRef), text ranges (PlcfandTxt), annotation bookmarks, authors, Word 2002+ extras */
  fcPlcfandRef: number
  lcbPlcfandRef: number
  fcPlcfandTxt: number
  lcbPlcfandTxt: number
  fcSttbfAtnBkmk: number
  lcbSttbfAtnBkmk: number
  fcPlcfAtnBkf: number
  lcbPlcfAtnBkf: number
  fcPlcfAtnBkl: number
  lcbPlcfAtnBkl: number
  fcSttbfRMark: number
  lcbSttbfRMark: number
  fcAtrdExtra: number
  lcbAtrdExtra: number
  /** text box stories of the body (PlcftxbxTxt) and the header document (PlcfHdrtxbxTxt) */
  fcPlcftxbxTxt: number
  lcbPlcftxbxTxt: number
  fcPlcfHdrtxbxTxt: number
  lcbPlcfHdrtxbxTxt: number
}

/** FibRgLw97 dword indices (the block starts right after the cslw count) */
const LW_CCP_TEXT = 3
const LW_CCP_FTN = 4
const LW_CCP_HDD = 5
const LW_CCP_ATN = 7
const LW_CCP_EDN = 8
const LW_CCP_TXBX = 9
const LW_CCP_HDR_TXBX = 10

/** FibRgFcLcb97 pair indices (MS-DOC 2.5.2 / 2.9.3xx) */
const PAIR_STSHF = 1
const PAIR_PLCFFND_REF = 2
const PAIR_PLCFFND_TXT = 3
const PAIR_PLCFAND_REF = 4
const PAIR_PLCFAND_TXT = 5
const PAIR_PLCFSED = 6
const PAIR_PLCFHDD = 11
const PAIR_PLCFBTE_CHPX = 12
const PAIR_STTBF_FFN = 15
const PAIR_PLFLST = 73
const PAIR_PLFLFO = 74
const PAIR_PLCFBTE_PAPX = 13
const PAIR_DOP = 31
const PAIR_CLX = 33
const PAIR_PLC_SPA_MOM = 40
const PAIR_PLC_SPA_HDR = 41
const PAIR_PLCFEND_REF = 46
const PAIR_PLCFEND_TXT = 47
const PAIR_STTBF_ATN_BKMK = 37
const PAIR_PLCF_ATN_BKF = 42
const PAIR_PLCF_ATN_BKL = 43
const PAIR_DGG_INFO = 50
const PAIR_STTBF_RMARK = 51
/** FibRgFcLcb2002: ATRDPost10 array (comment dates) */
const PAIR_ATRD_EXTRA = 112
const PAIR_PLCF_TXBX_TXT = 56
const PAIR_PLCF_HDR_TXBX_TXT = 58
/** Word 97's FibRgFcLcb97 has 93 pair entries; the CLX pair is inside every version */
const MIN_PAIR_COUNT = 34

function badFib(detail: string): never {
  throw new LegacyDocError('bad_fib', detail)
}

/** bounds-checked little-endian readers over the FIB bytes */
class FibReader {
  readonly view: DataView

  constructor(private readonly bytes: Uint8Array) {
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  }

  get length(): number {
    return this.bytes.length
  }

  need(offset: number, size: number, what: string): void {
    if (offset < 0 || size < 0 || offset + size > this.bytes.length)
      badFib(`the header ends before ${what} (needs bytes ${offset}-${offset + size})`)
  }

  u16(offset: number, what: string): number {
    this.need(offset, 2, what)
    return this.view.getUint16(offset, true)
  }

  u32(offset: number, what: string): number {
    this.need(offset, 4, what)
    return this.view.getUint32(offset, true)
  }
}

function fibPair(reader: FibReader, pairsStart: number, pairCount: number, index: number): FibPair {
  if (index >= pairCount)
    badFib(`the header has ${pairCount} fc/lcb pairs, so pair ${index} is missing`)
  const offset = pairsStart + index * 8
  return {
    fc: reader.u32(offset, `fc/lcb pair ${index}`),
    lcb: reader.u32(offset + 4, `fc/lcb pair ${index}`),
  }
}

function validatePairRange(pair: FibPair, streamLength: number, what: string): void {
  if (pair.lcb === 0) return
  if (pair.fc > streamLength || pair.lcb > streamLength - pair.fc)
    badFib(
      `the ${what} range (${pair.fc}+${pair.lcb}) lies outside the ${streamLength}-byte table stream`,
    )
}

/**
 * Parse the Word FIB (File Information Block) at the start of the
 * WordDocument stream: FibBase, FibRgW97, FibRgLw97 and FibRgFcLcb97.
 */
export function parseFib(wordDocument: Uint8Array): FibInfo {
  const reader = new FibReader(wordDocument)
  if (wordDocument.length < FIB_BASE_BYTES)
    badFib(
      `the WordDocument stream is only ${wordDocument.length} bytes, shorter than the 32-byte FibBase`,
    )

  const wIdent = reader.u16(0, 'FibBase.wIdent')
  if (wIdent !== FIB_IDENT)
    badFib(`FibBase.wIdent is 0x${wIdent.toString(16)}, expected 0x${FIB_IDENT.toString(16)}`)
  const nFib = reader.u16(2, 'FibBase.nFib')
  if (nFib < MIN_SUPPORTED_NFIB)
    throw new LegacyDocError(
      'unsupported_version',
      `nFib 0x${nFib.toString(16)} is a Word 6/95 or older document; Word 97-2003 is the oldest supported format`,
    )
  if (nFib > MAX_SUPPORTED_NFIB)
    throw new LegacyDocError(
      'unsupported_version',
      `nFib 0x${nFib.toString(16)} is newer than Word 2003`,
    )
  const flags = reader.u16(10, 'FibBase.flags')
  const encrypted = (flags & FLAG_ENCRYPTED) !== 0
  const obfuscated = (flags & FLAG_OBFUSCATED) !== 0
  if (encrypted) throw new LegacyDocError('encrypted', 'FibBase.fEncrypted is set')
  const tableStream = (flags & FLAG_WHICH_TABLE_STREAM) !== 0 ? '1Table' : '0Table'

  // FibRgW97: csw word count, then csw 2-byte values
  let offset = FIB_BASE_BYTES
  const csw = reader.u16(offset, 'csw')
  offset += 2
  reader.need(offset, csw * 2, 'FibRgW97')
  offset += csw * 2

  // FibRgLw97: cslw dword count, then cslw 4-byte values
  const cslw = reader.u16(offset, 'cslw')
  offset += 2
  if (cslw < LW_CCP_TEXT + 1) badFib(`FibRgLw97 has ${cslw} dwords, too few to carry ccpText`)
  reader.need(offset, cslw * 4, 'FibRgLw97')
  const lwStart = offset
  const lw = (index: number): number =>
    index < cslw ? reader.view.getUint32(lwStart + index * 4, true) : 0
  const ccpText = lw(LW_CCP_TEXT)
  const ccpFtn = lw(LW_CCP_FTN)
  const ccpHdd = lw(LW_CCP_HDD)
  const ccpAtn = lw(LW_CCP_ATN)
  const ccpEdn = lw(LW_CCP_EDN)
  const ccpTxbx = lw(LW_CCP_TXBX)
  const ccpHdrTxbx = lw(LW_CCP_HDR_TXBX)
  offset += cslw * 4

  // FibRgFcLcb97: pair count, then that many fc/lcb pairs
  const pairCount = reader.u16(offset, 'cbRgFcLcb')
  offset += 2
  if (pairCount < MIN_PAIR_COUNT)
    badFib(
      `the header has ${pairCount} fc/lcb pairs, fewer than the ${MIN_PAIR_COUNT} Word 97 needs`,
    )
  reader.need(offset, pairCount * 8, 'FibRgFcLcb97')
  const pairsStart = offset

  const stshf = fibPair(reader, pairsStart, pairCount, PAIR_STSHF)
  const plcfsed = fibPair(reader, pairsStart, pairCount, PAIR_PLCFSED)
  const plcfhdd = fibPair(reader, pairsStart, pairCount, PAIR_PLCFHDD)
  const plcffndRef = fibPair(reader, pairsStart, pairCount, PAIR_PLCFFND_REF)
  const plcffndTxt = fibPair(reader, pairsStart, pairCount, PAIR_PLCFFND_TXT)
  // lists are Word 97+ pairs 73/74; shorter synthetic headers simply have none
  const optionalPair = (index: number): FibPair =>
    index < pairCount ? fibPair(reader, pairsStart, pairCount, index) : { fc: 0, lcb: 0 }
  const plflst = optionalPair(PAIR_PLFLST)
  const plflfo = optionalPair(PAIR_PLFLFO)
  const plcSpaMom = optionalPair(PAIR_PLC_SPA_MOM)
  const plcSpaHdr = optionalPair(PAIR_PLC_SPA_HDR)
  const dggInfo = optionalPair(PAIR_DGG_INFO)
  const plcfendRef = optionalPair(PAIR_PLCFEND_REF)
  const plcfendTxt = optionalPair(PAIR_PLCFEND_TXT)
  const plcfandRef = fibPair(reader, pairsStart, pairCount, PAIR_PLCFAND_REF)
  const plcfandTxt = fibPair(reader, pairsStart, pairCount, PAIR_PLCFAND_TXT)
  const sttbfAtnBkmk = optionalPair(PAIR_STTBF_ATN_BKMK)
  const plcfAtnBkf = optionalPair(PAIR_PLCF_ATN_BKF)
  const plcfAtnBkl = optionalPair(PAIR_PLCF_ATN_BKL)
  const sttbfRMark = optionalPair(PAIR_STTBF_RMARK)
  const atrdExtra = optionalPair(PAIR_ATRD_EXTRA)
  const txbxTxt = optionalPair(PAIR_PLCF_TXBX_TXT)
  const hdrTxbxTxt = optionalPair(PAIR_PLCF_HDR_TXBX_TXT)
  const sttbfFfn = fibPair(reader, pairsStart, pairCount, PAIR_STTBF_FFN)
  const chpx = fibPair(reader, pairsStart, pairCount, PAIR_PLCFBTE_CHPX)
  const papx = fibPair(reader, pairsStart, pairCount, PAIR_PLCFBTE_PAPX)
  const dop = fibPair(reader, pairsStart, pairCount, PAIR_DOP)
  const clx = fibPair(reader, pairsStart, pairCount, PAIR_CLX)
  if (clx.lcb === 0) badFib('the header carries no CLX piece table (fcClx/lcbClx are empty)')

  return {
    nFib,
    flags,
    encrypted,
    obfuscated,
    tableStream,
    ccpText,
    ccpFtn,
    ccpHdd,
    ccpAtn,
    ccpEdn,
    ccpTxbx,
    ccpHdrTxbx,
    fcClx: clx.fc,
    lcbClx: clx.lcb,
    fcStshf: stshf.fc,
    lcbStshf: stshf.lcb,
    fcPlcfsed: plcfsed.fc,
    lcbPlcfsed: plcfsed.lcb,
    fcPlcfhdd: plcfhdd.fc,
    lcbPlcfhdd: plcfhdd.lcb,
    fcPlcffndRef: plcffndRef.fc,
    lcbPlcffndRef: plcffndRef.lcb,
    fcPlcffndTxt: plcffndTxt.fc,
    lcbPlcffndTxt: plcffndTxt.lcb,
    fcPlcfendRef: plcfendRef.fc,
    lcbPlcfendRef: plcfendRef.lcb,
    fcPlcfendTxt: plcfendTxt.fc,
    lcbPlcfendTxt: plcfendTxt.lcb,
    fcPlfLst: plflst.fc,
    lcbPlfLst: plflst.lcb,
    fcPlfLfo: plflfo.fc,
    lcbPlfLfo: plflfo.lcb,
    fcSttbfFfn: sttbfFfn.fc,
    lcbSttbfFfn: sttbfFfn.lcb,
    fcPlcfbteChpx: chpx.fc,
    lcbPlcfbteChpx: chpx.lcb,
    fcPlcfbtePapx: papx.fc,
    lcbPlcfbtePapx: papx.lcb,
    fcDop: dop.fc,
    lcbDop: dop.lcb,
    fcPlcSpaMom: plcSpaMom.fc,
    lcbPlcSpaMom: plcSpaMom.lcb,
    fcPlcSpaHdr: plcSpaHdr.fc,
    lcbPlcSpaHdr: plcSpaHdr.lcb,
    fcDggInfo: dggInfo.fc,
    lcbDggInfo: dggInfo.lcb,
    fcPlcfandRef: plcfandRef.fc,
    lcbPlcfandRef: plcfandRef.lcb,
    fcPlcfandTxt: plcfandTxt.fc,
    lcbPlcfandTxt: plcfandTxt.lcb,
    fcSttbfAtnBkmk: sttbfAtnBkmk.fc,
    lcbSttbfAtnBkmk: sttbfAtnBkmk.lcb,
    fcPlcfAtnBkf: plcfAtnBkf.fc,
    lcbPlcfAtnBkf: plcfAtnBkf.lcb,
    fcPlcfAtnBkl: plcfAtnBkl.fc,
    lcbPlcfAtnBkl: plcfAtnBkl.lcb,
    fcSttbfRMark: sttbfRMark.fc,
    lcbSttbfRMark: sttbfRMark.lcb,
    fcAtrdExtra: atrdExtra.fc,
    lcbAtrdExtra: atrdExtra.lcb,
    fcPlcftxbxTxt: txbxTxt.fc,
    lcbPlcftxbxTxt: txbxTxt.lcb,
    fcPlcfHdrtxbxTxt: hdrTxbxTxt.fc,
    lcbPlcfHdrtxbxTxt: hdrTxbxTxt.lcb,
  }
}

/**
 * Check the FibRgFcLcb97 ranges stages 2-3 will read against the table stream
 * they point into (all fc/lcb pairs are table-stream offsets, except the CLX
 * which clx.ts validates itself so it can report bad_clx).
 */
export function validateFibTableRanges(fib: FibInfo, tableLength: number): void {
  validatePairRange({ fc: fib.fcStshf, lcb: fib.lcbStshf }, tableLength, 'stylesheet (fcStshf)')
  validatePairRange(
    { fc: fib.fcPlcfsed, lcb: fib.lcbPlcfsed },
    tableLength,
    'section table (fcPlcfsed)',
  )
  validatePairRange(
    { fc: fib.fcPlcfhdd, lcb: fib.lcbPlcfhdd },
    tableLength,
    'header/footer table (fcPlcfhdd)',
  )
  for (const [fc, lcb, what] of [
    [fib.fcPlcffndRef, fib.lcbPlcffndRef, 'footnote reference table (fcPlcffndRef)'],
    [fib.fcPlcffndTxt, fib.lcbPlcffndTxt, 'footnote text table (fcPlcffndTxt)'],
    [fib.fcPlcfendRef, fib.lcbPlcfendRef, 'endnote reference table (fcPlcfendRef)'],
    [fib.fcPlcfendTxt, fib.lcbPlcfendTxt, 'endnote text table (fcPlcfendTxt)'],
  ] as const)
    validatePairRange({ fc, lcb }, tableLength, what)
  validatePairRange({ fc: fib.fcPlfLst, lcb: fib.lcbPlfLst }, tableLength, 'list table (fcPlfLst)')
  validatePairRange(
    { fc: fib.fcPlfLfo, lcb: fib.lcbPlfLfo },
    tableLength,
    'list override table (fcPlfLfo)',
  )
  validatePairRange(
    { fc: fib.fcSttbfFfn, lcb: fib.lcbSttbfFfn },
    tableLength,
    'font table (fcSttbfFfn)',
  )
  validatePairRange(
    { fc: fib.fcPlcfbteChpx, lcb: fib.lcbPlcfbteChpx },
    tableLength,
    'CHPX bin table (fcPlcfbteChpx)',
  )
  validatePairRange(
    { fc: fib.fcPlcfbtePapx, lcb: fib.lcbPlcfbtePapx },
    tableLength,
    'PAPX bin table (fcPlcfbtePapx)',
  )
  validatePairRange({ fc: fib.fcDop, lcb: fib.lcbDop }, tableLength, 'document properties (fcDop)')
}
