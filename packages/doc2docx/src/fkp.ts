import { LegacyDocError } from './errors'

/** every formatting FKP is one 512-byte page of the WordDocument stream */
export const FKP_PAGE_BYTES = 512
/** the crun byte is the last byte of the page */
const CRUN_OFFSET = FKP_PAGE_BYTES - 1
/** CHPX FKP layout: rgfc (4 bytes/run) + rgb (1 byte/run) + the crun byte: 5*crun + 4 <= 511 */
const MAX_CHPX_RUNS = 101
/** PAPX FKP layout: rgfc (4 bytes/run) + BxPap (13 bytes/run) + the crun byte: 17*crun + 4 <= 511 */
const MAX_PAPX_RUNS = 29
const BX_PAP_BYTES = 13

/** One CHPX run: the character formatting applied to an FC range. */
export interface ChpxRun {
  fcStart: number
  fcEnd: number
  /** the run's grpprl (sprms); empty when the run has no explicit formatting */
  grpprl: Uint8Array
}

/** One paragraph entry of a PAPX FKP: the paragraph's start FC and its PAPX. */
export interface PapxEntry {
  fcStart: number
  fcEnd: number
  /** paragraph style index (istd) from the PapxInFkp; 0 for a default paragraph */
  istd: number
  /** sprms after the istd; empty when the paragraph has no explicit properties */
  grpprl: Uint8Array
  /** true when BxPap.bOffset was 0: the FKP carries no PAPX for this paragraph */
  defaults: boolean
}

const EMPTY = new Uint8Array(0)

function badFkp(detail: string): never {
  throw new LegacyDocError('bad_fkp', detail)
}

function pageToOffset(wordDocument: Uint8Array, page: number, label: string): number {
  if (!Number.isInteger(page) || page < 0 || page > 0x3fffff)
    badFkp(`the ${label} FKP page number ${page} is invalid`)
  const base = page * FKP_PAGE_BYTES
  if (base + FKP_PAGE_BYTES > wordDocument.length)
    badFkp(
      `the ${label} FKP page ${page} (bytes ${base}-${base + FKP_PAGE_BYTES}) lies outside the ${wordDocument.length}-byte WordDocument stream`,
    )
  return base
}

/** the n+1 FC boundaries and the run count of one FKP page, with layout/order checks */
function readFkpHeader(
  wordDocument: Uint8Array,
  base: number,
  page: number,
  label: 'CHPX' | 'PAPX',
  maxRuns: number,
): { view: DataView; crun: number; fcs: number[]; dataStart: number } {
  const view = new DataView(wordDocument.buffer, wordDocument.byteOffset, wordDocument.byteLength)
  const crun = wordDocument[base + CRUN_OFFSET]!
  if (crun > maxRuns)
    badFkp(
      `the ${label} FKP page ${page} has ${crun} runs, more than the ${maxRuns} a 512-byte page can hold`,
    )
  const fcs: number[] = []
  for (let index = 0; index <= crun; index++) fcs.push(view.getUint32(base + index * 4, true))
  for (let index = 0; index < crun; index++) {
    if (fcs[index + 1]! < fcs[index]!)
      badFkp(`the ${label} FKP page ${page} has unordered FCs (${fcs[index]} > ${fcs[index + 1]})`)
  }
  const dataStart = base + 4 * (crun + 1) + (label === 'CHPX' ? crun : crun * BX_PAP_BYTES)
  if (dataStart > base + CRUN_OFFSET)
    badFkp(`the ${label} FKP page ${page} has no room for its run data`)
  return { view, crun, fcs, dataStart }
}

/**
 * Read one CHPX FKP page (MS-DOC 2.9.31): crun at byte 511, crun+1 FC
 * boundaries, crun byte offsets (offset/2) to the CHPX records. An offset of 0
 * means the run carries no explicit character properties.
 */
export function readChpxFkp(wordDocument: Uint8Array, page: number): ChpxRun[] {
  const base = pageToOffset(wordDocument, page, 'CHPX')
  const { crun, fcs, dataStart } = readFkpHeader(wordDocument, base, page, 'CHPX', MAX_CHPX_RUNS)
  const rgbStart = base + 4 * (crun + 1)
  const runs: ChpxRun[] = []
  for (let index = 0; index < crun; index++) {
    const offset = wordDocument[rgbStart + index]!
    let grpprl: Uint8Array = EMPTY
    if (offset !== 0) {
      const at = base + offset * 2
      if (at < dataStart || at >= base + CRUN_OFFSET)
        badFkp(`the CHPX of run ${index} on page ${page} points outside the run data area`)
      const cb = wordDocument[at]!
      if (at + 1 + cb > base + CRUN_OFFSET)
        badFkp(`the CHPX of run ${index} on page ${page} overruns the page`)
      grpprl = wordDocument.slice(at + 1, at + 1 + cb)
    }
    runs.push({ fcStart: fcs[index]!, fcEnd: fcs[index + 1]!, grpprl })
  }
  return runs
}

/**
 * Read one PAPX FKP page (MS-DOC 2.9.210): crun at byte 511, crun+1 paragraph
 * start FCs, crun BxPap records (bOffset + 12 reserved bytes). A bOffset of 0
 * means the paragraph has no PAPX (default properties). PapxInFkp starts with a
 * cb byte: cb != 0 stores a grpprl of 2*cb-1 bytes after it; cb == 0 is the
 * long form, where the next byte cb' stores 2*cb' bytes. The grpprl begins with
 * the 2-byte paragraph style index (istd).
 */
export function readPapxFkp(wordDocument: Uint8Array, page: number): PapxEntry[] {
  const base = pageToOffset(wordDocument, page, 'PAPX')
  const { crun, fcs, dataStart } = readFkpHeader(wordDocument, base, page, 'PAPX', MAX_PAPX_RUNS)
  const rgbxStart = base + 4 * (crun + 1)
  const entries: PapxEntry[] = []
  for (let index = 0; index < crun; index++) {
    const bOffset = wordDocument[rgbxStart + index * BX_PAP_BYTES]!
    let istd = 0
    let grpprl: Uint8Array = EMPTY
    const defaults = bOffset === 0
    if (!defaults) {
      const at = base + bOffset * 2
      if (at < dataStart || at >= base + CRUN_OFFSET)
        badFkp(`the PAPX of paragraph ${index} on page ${page} points outside the run data area`)
      const cb = wordDocument[at]!
      let raw: Uint8Array
      if (cb !== 0) {
        if (at + 1 + (2 * cb - 1) > base + CRUN_OFFSET)
          badFkp(`the PAPX of paragraph ${index} on page ${page} overruns the page`)
        raw = wordDocument.slice(at + 1, at + 1 + (2 * cb - 1))
      } else {
        if (at + 2 > base + CRUN_OFFSET)
          badFkp(`the PAPX of paragraph ${index} on page ${page} has no long-form length byte`)
        const longCb = wordDocument[at + 1]!
        if (at + 2 + 2 * longCb > base + CRUN_OFFSET)
          badFkp(`the PAPX of paragraph ${index} on page ${page} overruns the page`)
        raw = wordDocument.slice(at + 2, at + 2 + 2 * longCb)
      }
      if (raw.length < 2)
        badFkp(`the PAPX of paragraph ${index} on page ${page} is too short for its istd`)
      istd = raw[0]! | (raw[1]! << 8)
      grpprl = raw.subarray(2)
    }
    entries.push({ fcStart: fcs[index]!, fcEnd: fcs[index + 1]!, istd, grpprl, defaults })
  }
  return entries
}
