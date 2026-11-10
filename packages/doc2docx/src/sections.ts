/**
 * Stage 3c-1: section geometry from the PlcfSed table + SEPX section sprms.
 *
 * PlcfSed (FIB pair 6) is a PLC whose data entries are 12-byte SEDs:
 * fn u16, fcSepx u32, fnMpr u16, fcMpr u32. The CP array delimits the section
 * text ranges; fcSepx == 0xFFFFFFFF means the section uses the default SEP.
 * SEPX = u16 cbGrpprl + section sprms at fcSepx in the WordDocument stream.
 * Section sprms are matched by their low 9-bit operation (POI/MS-DOC section
 * sprm table), because producers vary the high spra bits (e.g. sprmSDyaTop is
 * 0x9023 in LibreOffice's output and 0xB023 in some references).
 *
 * Every section starts from Word's default section and then applies its own
 * sprms (0x80 keeps that default, 0x81 inverts a boolean). Documents without a
 * section table keep the blank template's section (the writer then leaves its
 * sectPr alone).
 */
import { LegacyDocError, isLegacyDocError } from './errors'
import type { FibInfo } from './fib'
import { iterateSprms, readSprmU16, readSprmU32, readSprmU8 } from './sprm'
import { borderLineAt, type BorderLine } from './tables'

export interface SectionMargins {
  left: number
  right: number
  top: number
  bottom: number
  header: number
  footer: number
  gutter: number
}

export type SectionBreakType = 'continuous' | 'column' | 'nextPage' | 'evenPage' | 'oddPage'

export type PageNumberFormat =
  'decimal' | 'upperRoman' | 'lowerRoman' | 'upperLetter' | 'lowerLetter'

export interface SectionColumns {
  count: number
  /** sprmSDxaColumns: default gap between columns */
  spaceTwips: number
  /** sprmSFLBetween: vertical rule between columns */
  separator: boolean
  /** sprmSDxaColWidth per column (unequal columns, sprmSFEvenlySpaced 0) */
  widths?: number[]
  /** sprmSDxaColSpacing per column (gap after that column) */
  spacings?: number[]
}

export type PageBorderSide = 'top' | 'left' | 'bottom' | 'right'

export interface PageBorderLine extends BorderLine {
  /** gap between the line and the text or page edge, in points (dptSpace) */
  spacePoints: number
}

export interface PageBorders {
  /** sprmSPgbProp pgbOffsetFrom: distance measured from the text or from the page edge */
  offsetFrom: 'text' | 'page'
  /** sprmSPgbProp pgbApplyTo; absent = all pages */
  display?: 'firstPage' | 'notFirstPage'
  /** sprmSPgbProp pgbPageDepth 1: the border is drawn behind the text */
  zOrder?: 'back'
  sides: Partial<Record<PageBorderSide, PageBorderLine>>
}

export interface SectionInfo {
  /** first CP of the section in the document text (main text CPs) */
  startCp: number
  /** one past the last CP of the section's text range */
  endCp: number
  pageWidthTwips: number
  pageHeightTwips: number
  margins: SectionMargins
  orientation: 'portrait' | 'landscape'
  titlePage: boolean
  /** sprmSBkc: how the section starts (w:type); absent = nextPage */
  breakType?: SectionBreakType
  /** text columns when the SEPX declares more than one */
  columns?: SectionColumns
  /** sprmSNfcPgn; absent = decimal */
  pageNumberFormat?: PageNumberFormat
  /** sprmSPgnStart, kept only when sprmSFPgnRestart is set */
  pageNumberStart?: number
  /** document grid (sprmSClm / sprmSDyaLinePitch / sprmSDxtCharSpace); no type = pitch kept, no snapping */
  docGrid?: {
    type?: 'lines' | 'linesAndChars' | 'snapToChars'
    linePitchTwips: number
    charSpace?: number
  }
  /** page borders (sprmSBrc*80 / sprmSBrc* + sprmSPgbProp); absent when no side has a line */
  pageBorders?: PageBorders
  /** non-fatal problems, e.g. multi-column sections that are flattened */
  warnings: string[]
}

export interface SectionSource {
  fib: FibInfo
  wordDocument: Uint8Array
  table: Uint8Array
}

/**
 * Fallback geometry when a SEPX is absent: mirrors docx-engine's blank
 * template sectPr (A4 11906x16838, 1440 twip margins, 708 header/footer).
 */
export const DEFAULT_SECTION_GEOMETRY = {
  pageWidthTwips: 11906,
  pageHeightTwips: 16838,
  margins: {
    left: 1440,
    right: 1440,
    top: 1440,
    bottom: 1440,
    header: 708,
    footer: 708,
    gutter: 0,
  },
  orientation: 'portrait' as const,
  titlePage: false,
}

const SEPX_DEFAULT = 0xffffffff
const PLCFSED_ENTRY_BYTES = 12
const PLCFSED_HEADER_BYTES = 4

function badSection(detail: string): never {
  throw new LegacyDocError('bad_section', detail)
}

function u16(bytes: Uint8Array, offset: number): number {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint16(offset, true)
}

function u32(bytes: Uint8Array, offset: number): number {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(offset, true)
}

/** section sprm operations (opcode & 0x1FF), MS-DOC 2.6.1 / POI SectionSprmUncompressor */
const SPRM_S_DXA_COL_WIDTH = 0x03
const SPRM_S_DXA_COL_SPACING = 0x04
const SPRM_S_F_EVENLY_SPACED = 0x05
const SPRM_S_F_L_BETWEEN = 0x06
const SPRM_S_BKC = 0x09
const SPRM_S_FTITLE_PAGE = 0x0a
const SPRM_S_CCOL_M1 = 0x0b
const SPRM_S_DXA_COLUMNS = 0x0c
const SPRM_S_NFC_PGN = 0x0e
const SPRM_S_F_PGN_RESTART = 0x11
const SPRM_S_PGN_START = 0x1c
const SPRM_S_DYA_HDR_TOP = 0x17
const SPRM_S_DYA_HDR_BOTTOM = 0x18
const SPRM_S_DM_ORIENT_PAGE = 0x1d
const SPRM_S_XA_PAGE = 0x1f
const SPRM_S_YA_PAGE = 0x20
const SPRM_S_DXA_LEFT = 0x21
const SPRM_S_DXA_RIGHT = 0x22
const SPRM_S_DYA_TOP = 0x23
const SPRM_S_DYA_BOTTOM = 0x24
const SPRM_S_DZA_GUTTER = 0x25
const SPRM_S_BRC_TOP_80 = 0x2b
const SPRM_S_BRC_LEFT_80 = 0x2c
const SPRM_S_BRC_BOTTOM_80 = 0x2d
const SPRM_S_BRC_RIGHT_80 = 0x2e
const SPRM_S_PGB_PROP = 0x2f
const SPRM_S_BRC_TOP = 0x34
const SPRM_S_BRC_LEFT = 0x35
const SPRM_S_BRC_BOTTOM = 0x36
const SPRM_S_BRC_RIGHT = 0x37
const PAGE_BORDER_SIDES: Record<number, PageBorderSide> = {
  [SPRM_S_BRC_TOP_80]: 'top',
  [SPRM_S_BRC_LEFT_80]: 'left',
  [SPRM_S_BRC_BOTTOM_80]: 'bottom',
  [SPRM_S_BRC_RIGHT_80]: 'right',
  [SPRM_S_BRC_TOP]: 'top',
  [SPRM_S_BRC_LEFT]: 'left',
  [SPRM_S_BRC_BOTTOM]: 'bottom',
  [SPRM_S_BRC_RIGHT]: 'right',
}
const SPRM_S_DXT_CHAR_SPACE = 0x30
const SPRM_S_DYA_LINE_PITCH = 0x31
const SPRM_S_CLM = 0x32
const BREAK_TYPES: Record<number, SectionBreakType> = {
  0: 'continuous',
  1: 'column',
  2: 'nextPage',
  3: 'evenPage',
  4: 'oddPage',
}
const PAGE_NUMBER_FORMATS: Record<number, PageNumberFormat> = {
  0: 'decimal',
  1: 'upperRoman',
  2: 'lowerRoman',
  3: 'upperLetter',
  4: 'lowerLetter',
}
const GRID_TYPES: Record<number, 'lines' | 'linesAndChars' | 'snapToChars'> = {
  1: 'linesAndChars',
  2: 'lines',
  3: 'snapToChars',
}

/** boolean sprm operand: 0 = off, 1 = on, 0x80 = inherit, 0x81 = invert */
function applyBoolean(
  current: boolean,
  value: number,
  mode: 'set' | 'invert' | 'inherit',
): boolean {
  if (mode === 'inherit') return current
  if (mode === 'invert') return !current
  return value !== 0
}

type SectionDraft = Omit<SectionInfo, 'startCp' | 'endCp' | 'warnings'>

/** a Brc80 (4 bytes) or Brc (8 bytes) page border; null = the side is explicitly off */
function pageBorderOf(operand: Uint8Array, wide: boolean): PageBorderLine | null | undefined {
  if (operand.length < (wide ? 8 : 4)) return undefined
  const line = borderLineAt(operand, 0, wide)
  if (!line || operand[wide ? 4 : 0] === 0) return null
  return { ...line, spacePoints: operand[wide ? 6 : 3]! & 0x1f }
}

/** sprmSPgbProp PGPInfo: pgbApplyTo bits 0-2, pgbPageDepth bits 3-4, pgbOffsetFrom bit 5 */
function applyPageBorderProps(borders: PageBorders, value: number): void {
  const applyTo = value & 0x7
  if (applyTo === 1) borders.display = 'firstPage'
  else if (applyTo === 2) borders.display = 'notFirstPage'
  else delete borders.display
  if (((value >> 3) & 0x3) === 1) borders.zOrder = 'back'
  else delete borders.zOrder
  borders.offsetFrom = value & 0x20 ? 'page' : 'text'
}

function applySepx(draft: SectionDraft, grpprl: Uint8Array, sectionIndex: number): string[] {
  const warnings: string[] = []
  let sprms
  try {
    sprms = iterateSprms(grpprl)
  } catch (error) {
    if (isLegacyDocError(error))
      badSection(`the SEPX of section ${sectionIndex} is invalid: ${error.message}`)
    throw error
  }
  // SEPX sprms come in ispmd order (char space, line pitch, then clm): assemble the grid after the walk
  let gridType: 'lines' | 'linesAndChars' | 'snapToChars' | null | undefined
  let linePitch: number | undefined
  let charSpace: number | undefined
  // the SEP defaults: one column, 720 twip gap, evenly spaced
  let columnCount = 1
  let columnSpace = 720
  let evenlySpaced = true
  let separator = false
  const columnWidths: number[] = []
  const columnSpacings: number[] = []
  let pageNumberStart: number | undefined
  let pageNumberRestart = false
  const pageBorders: PageBorders = draft.pageBorders
    ? { ...draft.pageBorders, sides: { ...draft.pageBorders.sides } }
    : { offsetFrom: 'text', sides: {} }
  for (const sprm of sprms) {
    const operation = sprm.opcode & 0x1ff
    const byte = readSprmU8(sprm.operand)
    const word = readSprmU16(sprm.operand)
    const borderSide = PAGE_BORDER_SIDES[operation]
    if (borderSide) {
      const line = pageBorderOf(sprm.operand, operation >= SPRM_S_BRC_TOP)
      if (line === null) delete pageBorders.sides[borderSide]
      else if (line) pageBorders.sides[borderSide] = line
      continue
    }
    switch (operation) {
      case SPRM_S_PGB_PROP:
        if (word !== undefined) applyPageBorderProps(pageBorders, word)
        break
      case SPRM_S_DYA_HDR_TOP:
        if (word !== undefined) draft.margins.header = word
        break
      case SPRM_S_DYA_HDR_BOTTOM:
        if (word !== undefined) draft.margins.footer = word
        break
      case SPRM_S_XA_PAGE:
        if (word !== undefined) draft.pageWidthTwips = word
        break
      case SPRM_S_YA_PAGE:
        if (word !== undefined) draft.pageHeightTwips = word
        break
      case SPRM_S_DXA_LEFT:
        if (word !== undefined) draft.margins.left = word
        break
      case SPRM_S_DXA_RIGHT:
        if (word !== undefined) draft.margins.right = word
        break
      case SPRM_S_DYA_TOP:
        if (word !== undefined) draft.margins.top = word
        break
      case SPRM_S_DYA_BOTTOM:
        if (word !== undefined) draft.margins.bottom = word
        break
      case SPRM_S_DZA_GUTTER:
        if (word !== undefined) draft.margins.gutter = word
        break
      case SPRM_S_BKC:
        if (byte !== undefined && BREAK_TYPES[byte]) draft.breakType = BREAK_TYPES[byte]
        break
      case SPRM_S_FTITLE_PAGE:
        if (byte !== undefined)
          draft.titlePage = applyBoolean(
            draft.titlePage,
            byte,
            byte === 0x80 ? 'inherit' : byte === 0x81 ? 'invert' : 'set',
          )
        break
      case SPRM_S_DM_ORIENT_PAGE:
        if (byte !== undefined)
          draft.orientation = applyBoolean(
            draft.orientation === 'landscape',
            byte,
            byte === 0x80 ? 'inherit' : byte === 0x81 ? 'invert' : 'set',
          )
            ? 'landscape'
            : 'portrait'
        break
      case SPRM_S_CCOL_M1:
        if (word !== undefined) columnCount = word + 1
        break
      case SPRM_S_DXA_COLUMNS:
        if (word !== undefined) columnSpace = word
        break
      case SPRM_S_F_EVENLY_SPACED:
        if (byte !== undefined) evenlySpaced = byte !== 0
        break
      case SPRM_S_F_L_BETWEEN:
        if (byte !== undefined) separator = byte !== 0
        break
      case SPRM_S_DXA_COL_WIDTH:
      case SPRM_S_DXA_COL_SPACING: {
        // operand: column index (u8) + twips (u16)
        const index = byte
        const value = readSprmU16(sprm.operand, 1)
        if (index === undefined || value === undefined) break
        ;(operation === SPRM_S_DXA_COL_WIDTH ? columnWidths : columnSpacings)[index] = value
        break
      }
      case SPRM_S_NFC_PGN:
        if (byte !== undefined && PAGE_NUMBER_FORMATS[byte])
          draft.pageNumberFormat = PAGE_NUMBER_FORMATS[byte]
        break
      case SPRM_S_PGN_START:
        if (word !== undefined) pageNumberStart = word
        break
      case SPRM_S_F_PGN_RESTART:
        if (byte !== undefined) pageNumberRestart = byte !== 0
        break
      case SPRM_S_CLM:
        if (word !== undefined) gridType = GRID_TYPES[word] ?? null
        break
      case SPRM_S_DYA_LINE_PITCH:
        if (word !== undefined) linePitch = word
        break
      case SPRM_S_DXT_CHAR_SPACE: {
        const value = readSprmU32(sprm.operand)
        if (value !== undefined) charSpace = value | 0
        break
      }
      default:
        // unknown/unsupported section sprm: skipped by the operand-size rule
        break
    }
  }
  if (columnCount > 1) {
    const widths = columnWidths.slice(0, columnCount)
    const explicit = !evenlySpaced && widths.length === columnCount && widths.every((w) => w > 0)
    draft.columns = {
      count: columnCount,
      spaceTwips: columnSpace,
      separator,
      ...(explicit ? { widths } : {}),
      ...(explicit && columnSpacings.length > 0
        ? {
            spacings: Array.from(
              { length: columnCount },
              (_, i) => columnSpacings[i] ?? columnSpace,
            ),
          }
        : {}),
    }
  } else delete draft.columns
  if (Object.keys(pageBorders.sides).length > 0) draft.pageBorders = pageBorders
  else delete draft.pageBorders
  if (pageNumberRestart) draft.pageNumberStart = pageNumberStart ?? 1
  if (gridType || linePitch !== undefined) {
    draft.docGrid = {
      ...(gridType ? { type: gridType } : {}),
      linePitchTwips: linePitch ?? 0,
      ...(charSpace && gridType && gridType !== 'lines' ? { charSpace } : {}),
    }
  } else delete draft.docGrid
  return warnings
}

/**
 * Parse the PlcfSed section table and each section's SEPX. Returns [] when the
 * document carries no section table (the writer then keeps the template sectPr).
 */
/** The end CPs of every section but the last (the section marks), or [] when the table is unusable. */
export function sectionMarkCps(fib: FibInfo, table: Uint8Array): number[] {
  const lcb = fib.lcbPlcfsed
  if (lcb < PLCFSED_HEADER_BYTES + PLCFSED_ENTRY_BYTES || (lcb - 4) % 16 !== 0) return []
  if (fib.fcPlcfsed > table.length || lcb > table.length - fib.fcPlcfsed) return []
  const count = (lcb - 4) / 16
  const cps: number[] = []
  for (let index = 1; index < count; index++) cps.push(u32(table, fib.fcPlcfsed + index * 4))
  return cps.filter((cp) => cp <= fib.ccpText)
}

export function parseSections(source: SectionSource): SectionInfo[] {
  const { fib, wordDocument, table } = source
  const lcb = fib.lcbPlcfsed
  if (lcb === 0) return []
  if (lcb < PLCFSED_HEADER_BYTES + PLCFSED_ENTRY_BYTES || (lcb - 4) % 16 !== 0)
    badSection(`the section table has an invalid length (lcb ${lcb}, expected 16n+4)`)
  if (fib.fcPlcfsed > table.length || lcb > table.length - fib.fcPlcfsed)
    badSection(
      `the section table range (${fib.fcPlcfsed}+${lcb}) lies outside the ${table.length}-byte table stream`,
    )
  const count = (lcb - 4) / 16
  const cps: number[] = []
  for (let index = 0; index <= count; index++) cps.push(u32(table, fib.fcPlcfsed + index * 4))
  if (cps[0] !== 0) badSection(`the section table does not start at CP 0 (starts at ${cps[0]})`)
  for (let index = 0; index < count; index++) {
    if (cps[index + 1]! < cps[index]!)
      badSection(`the section CPs are not ordered (${cps[index]} > ${cps[index + 1]})`)
  }
  if (cps[count]! < fib.ccpText)
    badSection(
      `the section table ends at CP ${cps[count]}, before the ${fib.ccpText} body characters`,
    )
  const sections: SectionInfo[] = []
  for (let index = 0; index < count; index++) {
    const draft: SectionDraft = {
      pageWidthTwips: DEFAULT_SECTION_GEOMETRY.pageWidthTwips,
      pageHeightTwips: DEFAULT_SECTION_GEOMETRY.pageHeightTwips,
      margins: { ...DEFAULT_SECTION_GEOMETRY.margins },
      orientation: DEFAULT_SECTION_GEOMETRY.orientation,
      titlePage: DEFAULT_SECTION_GEOMETRY.titlePage,
    }
    const sedOffset = fib.fcPlcfsed + (count + 1) * 4 + index * PLCFSED_ENTRY_BYTES
    const fcSepx = u32(table, sedOffset + 2)
    const warnings: string[] = []
    if (fcSepx !== SEPX_DEFAULT) {
      if (fcSepx + 2 > wordDocument.length)
        badSection(
          `the SEPX of section ${index} (fc ${fcSepx}) lies outside the WordDocument stream`,
        )
      const cb = u16(wordDocument, fcSepx)
      if (fcSepx + 2 + cb > wordDocument.length)
        badSection(`the SEPX of section ${index} overruns the WordDocument stream`)
      const grpprl = wordDocument.subarray(fcSepx + 2, fcSepx + 2 + cb)
      warnings.push(...applySepx(draft, grpprl, index))
    }
    sections.push({
      startCp: cps[index]!,
      endCp: cps[index + 1]!,
      ...draft,
      warnings: [...new Set(warnings)],
    })
  }
  return sections
}
