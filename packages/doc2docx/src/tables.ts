/**
 * Stage 3c-2: Word 97-2003 table reconstruction.
 *
 * The body text marks a cell with \x07 and the row end with an extra empty
 * \x07. Each row-end paragraph has a PAPX whose TAP sprms (sprm type 5) carry
 * the table definition: sprmTDefTable (op 0x08, a u16-length operand) holds
 * the per-row cell boundaries (rgdxaCenter, twips) and one 20-byte TC
 * descriptor per physical cell (merge flags, width, four BorderCodes), while
 * the other TAP sprms set alignment, cantSplit, header row, row height,
 * borders and cell padding.
 *
 * Cells become real w:tbl tables in the writer; cell character formatting is
 * resolved from the CHPX runs by CP, cell paragraph formatting stays plain
 * (with a warning when the PAPX carries explicit paragraph properties).
 * Nested tables are detected and reported, but only single-level tables are
 * reconstructed. Malformed TAP/TDefTable data degrades: the affected row (or,
 * when the grid cannot be built, the whole table) is imported as plain
 * paragraphs and the problem is reported in `warnings`.
 */
import {
  applyParagraphGrpprl,
  cpOfFc,
  expandHugePapx,
  readFormattingDetailed,
  type FormattingRead,
  type FormattedParagraph,
  type FormattingSource,
  type NoteMark,
  type ParagraphContext,
  type ParagraphFormatting,
  type ParagraphFrame,
  isPositionedFrame,
  type Shading,
  shading80At,
  shadingAt,
} from './formatting'
import { readPapxFkp } from './fkp'
import type { LegacyDocNotes } from './notes'
import { parsePlcBte } from './plcf'
import {
  commentMarksWithin,
  fieldRangesWithin,
  noteMarksWithin,
  markFormattingEntry,
  readStyleSheet,
  resolveParagraphFormattingFor,
  resolveParagraphMarkFormatting,
  resolveRunsForRange,
  type CommentMarkAt,
  type FormattedLinkRange,
  type ParagraphRunRequest,
} from './resolve'
import type { LegacyComment } from './comments'
import { iterateSprmsLenient, readSprmI16, readSprmU16, readSprmU8 } from './sprm'
import type { FieldParseResult, SimpleField } from './fields'
import type { ImageBlock } from './images'
import type { FloatingShape } from './drawings'
import { isBulletLevel, resolveListLevel, type ListTables } from './lists'
import type { StyleSheet } from './styles'
import { COLOR_BY_INDEX } from './formatting'
import { mapWordControlChars } from './text'

/* TAP sprm operations (opcode & 0x1FF; sgc 5) */
const SPRM_T_JC = 0x00
const SPRM_T_CANT_SPLIT = 0x03
const SPRM_T_TABLE_HEADER = 0x04
const SPRM_T_TABLE_BORDERS = 0x05
/** sprmTTableBorders (Word 2003+): six 8-byte Brc */
const SPRM_T_TABLE_BORDERS_WIDE = 0x13
/** sprmTSetBrc80 / sprmTSetBrc: itcFirst, itcLim, grfbrc, then a Brc80 / Brc */
const SPRM_T_SET_BRC80 = 0x20
const SPRM_T_SET_BRC = 0x2f
const SPRM_T_DYA_ROW_HEIGHT = 0x07
const SPRM_T_TDEF_TABLE = 0x08
const SPRM_T_PC = 0x0d
const SPRM_T_DXA_ABS = 0x0e
const SPRM_T_DYA_ABS = 0x0f
const SPRM_T_DXA_FROM_TEXT = 0x10
const SPRM_T_DYA_FROM_TEXT = 0x11
const SPRM_T_DXA_FROM_TEXT_RIGHT = 0x1e
const SPRM_T_DYA_FROM_TEXT_BOTTOM = 0x1f
const SPRM_T_F_NO_ALLOW_OVERLAP = 0x65
const POSITION_H_ANCHORS = ['text', 'margin', 'page'] as const
const POSITION_V_ANCHORS = ['margin', 'page', 'text'] as const
/** sprmTDxaGapHalf: half the gap between cells, i.e. the default left/right cell margin */
const SPRM_T_DXA_GAP_HALF = 0x02
/** sprmTCellPaddingDefault (whole row) and sprmTCellPadding (itcFirst..itcLim) */
const SPRM_T_CELL_PADDING_DEFAULT = 0x34
const SPRM_T_CELL_PADDING = 0x32
/** sprmTVertAlign: itcFirst, itcLim, vertAlign (0 top, 1 center, 2 bottom) */
const SPRM_T_VERT_ALIGN = 0x2c
/** sprmTDefTableShd80: one Shd80 per cell */
const SPRM_T_DEF_TABLE_SHD80 = 0x09
/** sprmTDefTableShd / 2nd / 3rd: one Shd per cell for cells 0-21 / 22-43 / 44-62 */
const SPRM_T_DEF_TABLE_SHD: Record<number, number> = { 0x12: 0, 0x16: 22, 0x0c: 44 }
/** sprmTDefTableShdRaw / 2nd / 3rd: the same cell shading before table-style resolution */
const SPRM_T_DEF_TABLE_SHD_RAW: Record<number, number> = { 0x70: 0, 0x71: 22, 0x72: 44 }
/** sprmTSetShd / sprmTSetShdOdd: TableShadeOperand itcFirst, itcLim, Shd (odd = every other cell) */
const SPRM_T_SET_SHD = 0x2d
const SPRM_T_SET_SHD_ODD = 0x2e
/** sprmTSetShdTable: SHDOperand (cb, Shd) for the whole table */
const SPRM_T_SET_SHD_TABLE = 0x60
/** Word's default half gap, used when a row has no sprmTDxaGapHalf */
const DEFAULT_GAP_HALF = 108
/** sprmPItap / sprmPDtap: table nesting level (absolute / delta) stored in the paragraph PAPX */
const SPRM_P_ITAP = 0x6649
const SPRM_P_DTAP = 0x664a
/** sprmPFInnerTableCell / sprmPFInnerTtp: a paragraph mark that ends an inner cell / inner row */
const SPRM_P_F_INNER_TABLE_CELL = 0x244b
const SPRM_P_F_INNER_TTP = 0x244c
/** TC rgf merge bits (TCAbstractType): fMerged marks continuation cells */
const TC_MERGED = 0x0002
const TC_VERT_MERGE = 0x0020
const TC_VERT_RESTART = 0x0040
const TC_BYTES = 20
const MAX_GRID_COLUMNS = 63
const EMPTY = new Uint8Array(0)

export interface CellBorders {
  top: boolean
  left: boolean
  bottom: boolean
  right: boolean
}

export interface TableBorders {
  top: boolean
  left: boolean
  bottom: boolean
  right: boolean
  insideHorizontal: boolean
  insideVertical: boolean
}

/** one border line as OOXML writes it (w:val / w:sz in eighths of a point / w:color) */
export interface BorderLine {
  style: string
  widthEighths: number
  color?: string
}

export type BorderSide = keyof TableBorders

/** the per-side line styles behind a TableBorders / CellBorders set (absent side = default single line) */
export type BorderStyles = Partial<Record<BorderSide, BorderLine>>

export interface TableCellPadding {
  top?: number
  left?: number
  bottom?: number
  right?: number
}

export type CellVerticalAlign = 'center' | 'bottom'

export interface TableCell {
  /** number of grid columns this cell covers */
  gridSpan: number
  /** TC80 vertAlign bits / sprmTVertAlign; absent = top */
  vertAlign?: CellVerticalAlign
  shading?: Shading
  /** sprmTCellPadding: this cell's margins (w:tcMar), overriding the table defaults */
  padding?: TableCellPadding
  /** width in twips (sum of the covered grid columns) */
  widthTwips: number
  /** vertical merge state from the TC rgf bits */
  verticalMerge?: 'restart' | 'continue'
  /** visible borders of this cell (undefined when the TC carries none) */
  borders?: CellBorders
  borderStyles?: BorderStyles
  /** cell content; always at least one (possibly empty) paragraph */
  paragraphs: FormattedParagraph[]
  /** list membership of the cell paragraphs, by paragraph index */
  paragraphLists?: Map<number, ParagraphListInfo>
  /** pictures anchored in the cell paragraphs, by paragraph index */
  paragraphPictures?: Map<number, ParagraphPictures>
  /** nested tables, keyed by the index of the cell paragraph that follows them */
  nestedTables?: Map<number, TableModel[]>
}

export interface TableRow {
  cells: TableCell[]
  /** sprmTTableHeader: repeat this row at the top of each page */
  header: boolean
  /** sprmTFCantSplit: keep the row on one page */
  cantSplit: boolean
  /** sprmTDyaRowHeight in twips (absolute value) */
  heightTwips?: number
  /** true when the row height is an exact value (negative dyaRowHeight) */
  heightExact?: boolean
}

/** Floating table position (sprmTPc / sprmTDxaAbs / sprmTDyaAbs / sprmTDxaFromText ...), written as w:tblpPr. */
export interface TablePosition {
  /** twips, or the alignment codes -4 center, -8 right, -12 inside, -16 outside */
  xTwips?: number
  /** twips, or -4 top, -8 center, -12 bottom, -16 inside, -20 outside */
  yTwips?: number
  hAnchor?: 'text' | 'margin' | 'page'
  vAnchor?: 'text' | 'margin' | 'page'
  leftFromText?: number
  rightFromText?: number
  topFromText?: number
  bottomFromText?: number
  /** sprmTFNoAllowOverlap */
  noOverlap?: boolean
}

export interface TableModel {
  /** width of each grid column in twips (start at the first cell boundary) */
  gridWidths: number[]
  rows: TableRow[]
  alignment?: 'left' | 'center' | 'right'
  /** visible table borders */
  borders: TableBorders
  /** line styles of the visible borders (sprmTTableBorders); sides without one are single 0.5pt */
  borderStyles?: BorderStyles
  /** default cell margins (w:tblCellMar): left/right from sprmTDxaGapHalf, then sprmTCellPaddingDefault */
  cellPadding: TableCellPadding
  shading?: Shading
  position?: TablePosition
  warnings: string[]
}

/** List membership of a body paragraph (stage 3c-4); numId is 100 + ilfo. */
export interface ParagraphListInfo {
  kind: 'bullet' | 'ordered'
  numId: string
  ilvl: number
}

/** Pictures anchored inside a paragraph: floating shapes and inline pictures mixed with text. */
export interface ParagraphPictures {
  floating?: FloatingShape[]
  /** inline pictures that share the paragraph with text; placed at their run's pictureCp, else at the start or end */
  inline?: Array<{ image: ImageBlock; position: 'start' | 'end' }>
}

export type BodyBlock =
  | {
      type: 'paragraph'
      paragraph: FormattedParagraph
      list?: ParagraphListInfo
      pictures?: ParagraphPictures
    }
  | { type: 'table'; table: TableModel }
  | {
      type: 'image'
      image: ImageBlock
      /** the picture's own paragraph (alignment, spacing, indents) */
      paragraph?: FormattedParagraph
      list?: ParagraphListInfo
    }

/** The reconstructed body: paragraphs, tables and pictures in CP (document) order. */
export interface LegacyDocBody {
  blocks: BodyBlock[]
  warnings: string[]
  /** list definitions/overrides the numbered paragraphs refer to (stage 3c-4) */
  lists?: ListTables
  /** footnotes/endnotes referenced from the body (notes.ts) */
  notes?: LegacyDocNotes
  /** comments anchored in the body (their marks sit in the paragraph runs) */
  comments?: LegacyComment[]
}

/** One raw unit of the body text; markers keep their CP so formatting can be sliced. */
interface BodyUnit {
  kind: 'paragraph' | 'cell' | 'rowEnd'
  startCp: number
  endCp: number
  /** text between the previous mark (or CP 0) and this unit's mark */
  text: string
}

interface TapTableCell {
  rgf: number
  wWidth: number
  borders: CellBorders
  borderStyles: BorderStyles
  vertAlign?: CellVerticalAlign
  shading?: Shading
  /** set once a resolved shading sprm wrote this cell, so the Shd80/raw twins no longer apply */
  shadingResolved?: boolean
  padding?: TableCellPadding
}

function defaultTapCell(): TapTableCell {
  return {
    rgf: 0,
    wWidth: 0,
    borders: { top: false, left: false, bottom: false, right: false },
    borderStyles: {},
  }
}

interface TapTableDefinition {
  /** cell boundaries from rgdxaCenter, in twips, before the first boundary */
  centers: number[]
  cells: TapTableCell[]
}

interface TapProperties {
  table?: TapTableDefinition
  alignment?: 'left' | 'center' | 'right'
  cantSplit: boolean
  header: boolean
  heightTwips?: number
  heightExact?: boolean
  borders: TableBorders
  borderStyles: BorderStyles
  gapHalf?: number
  cellPadding?: TableCellPadding
  position?: TablePosition
  shading?: Shading
}

/** a structural problem of one row or table; the caller falls back to paragraphs */
class TableProblem extends Error {}

function badTable(detail: string): never {
  throw new TableProblem(detail)
}

function u16At(bytes: Uint8Array, offset: number): number {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint16(offset, true)
}

function i16At(bytes: Uint8Array, offset: number): number {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getInt16(offset, true)
}

/** a BRC is visible when its brcType is a real border style (not "none") */
function visibleBorderCode(border: Uint8Array, offset: number): boolean {
  const brcType = border[offset + 1]
  return brcType !== undefined && brcType !== 0 && brcType <= 0x1b
}

/** MS-DOC brcType -> OOXML ST_Border */
const BORDER_STYLES: Record<number, string> = {
  1: 'single',
  2: 'thick',
  3: 'double',
  5: 'hairline',
  6: 'dotted',
  7: 'dashed',
  8: 'dotDash',
  9: 'dotDotDash',
  10: 'triple',
  11: 'thinThickSmallGap',
  12: 'thickThinSmallGap',
  13: 'thinThickThinSmallGap',
  14: 'thinThickMediumGap',
  15: 'thickThinMediumGap',
  16: 'thinThickThinMediumGap',
  17: 'thinThickLargeGap',
  18: 'thickThinLargeGap',
  19: 'thinThickThinLargeGap',
  20: 'wave',
  21: 'doubleWave',
  22: 'dashSmallGap',
  23: 'dashDotStroked',
  24: 'threeDEmboss',
  25: 'threeDEngrave',
  26: 'outset',
  27: 'inset',
}

/** a Brc80 (4 bytes: width, type, ico, space) or Brc (8 bytes: cv, width, type, space); undefined = no line */
export function borderLineAt(
  bytes: Uint8Array,
  offset: number,
  wide: boolean,
): BorderLine | undefined {
  if (offset + (wide ? 8 : 4) > bytes.length) return undefined
  const base = wide ? offset + 4 : offset
  const widthEighths = bytes[base]!
  const brcType = bytes[base + 1]!
  if (brcType === 0 || brcType === 0xff) return undefined
  // Word 6 "thick" (type 2) draws the line at twice the stored width
  const width = brcType === 2 ? widthEighths * 2 : widthEighths
  const style = BORDER_STYLES[brcType] ?? 'single'
  let color: string | undefined
  if (wide) {
    if (bytes[offset + 3] !== 0xff)
      color = [bytes[offset]!, bytes[offset + 1]!, bytes[offset + 2]!]
        .map((channel) => channel.toString(16).padStart(2, '0'))
        .join('')
        .toUpperCase()
  } else color = COLOR_BY_INDEX[bytes[base + 2]!]
  return {
    style,
    widthEighths: width > 0 ? width : 4,
    ...(color !== undefined && color !== 'auto' ? { color } : {}),
  }
}

const TABLE_BORDER_SIDES: BorderSide[] = [
  'top',
  'left',
  'bottom',
  'right',
  'insideHorizontal',
  'insideVertical',
]

function tableBorderSet(
  operand: Uint8Array,
  wide: boolean,
): {
  borders: TableBorders
  styles: BorderStyles
} {
  const borders = emptyBorders()
  const styles: BorderStyles = {}
  TABLE_BORDER_SIDES.forEach((side, index) => {
    const line = borderLineAt(operand, index * (wide ? 8 : 4), wide)
    if (!line) return
    borders[side] = true
    styles[side] = line
  })
  return { borders, styles }
}

function emptyBorders(): TableBorders {
  return {
    top: false,
    left: false,
    bottom: false,
    right: false,
    insideHorizontal: false,
    insideVertical: false,
  }
}

/** toggle-style operand: 0 off, 1 on, 0x80 keep the inherited value, 0x81 invert it */
function toggleOperand(value: number | undefined, current: boolean): boolean {
  if (value === undefined) return current
  if (value === 0x80) return current
  if (value === 0x81) return !current
  return value !== 0
}

function parseTDefTable(operand: Uint8Array, rowIndex: number): TapTableDefinition {
  if (operand.length < 1) badTable('the TDefTable of row ' + rowIndex + ' is empty')
  const itcMac = operand[0]!
  if (itcMac < 1 || itcMac > MAX_GRID_COLUMNS)
    badTable('the TDefTable of row ' + rowIndex + ' declares ' + itcMac + ' cells')
  const tcStart = 1 + (itcMac + 1) * 2
  if (operand.length < tcStart)
    badTable(
      'the TDefTable of row ' +
        rowIndex +
        ' needs ' +
        tcStart +
        ' bytes for ' +
        itcMac +
        ' cell boundaries but has ' +
        operand.length,
    )
  const centers: number[] = []
  for (let index = 0; index <= itcMac; index++) centers.push(i16At(operand, 1 + index * 2))
  for (let index = 0; index < itcMac; index++) {
    if (centers[index + 1]! <= centers[index]!)
      badTable(
        'the TDefTable of row ' +
          rowIndex +
          ' has unordered cell boundaries (' +
          centers[index] +
          ' >= ' +
          centers[index + 1] +
          ')',
      )
  }
  const cells: TapTableCell[] = []
  for (let index = 0; index < itcMac; index++) {
    const base = tcStart + index * TC_BYTES
    // Word writes rows whose rgTc80 is truncated or absent; the missing cells take default TC80s
    if (base + TC_BYTES > operand.length) {
      cells.push(defaultTapCell())
      continue
    }
    const rgf = u16At(operand, base)
    const vertAlign = vertAlignOf((rgf >> 7) & 3)
    const borders: CellBorders = {
      top: visibleBorderCode(operand, base + 4),
      left: visibleBorderCode(operand, base + 8),
      bottom: visibleBorderCode(operand, base + 12),
      right: visibleBorderCode(operand, base + 16),
    }
    const borderStyles: BorderStyles = {}
    ;(['top', 'left', 'bottom', 'right'] as const).forEach((side, sideIndex) => {
      const line = borderLineAt(operand, base + 4 + sideIndex * 4, false)
      if (line && borders[side]) borderStyles[side] = line
    })
    cells.push({
      rgf,
      wWidth: u16At(operand, base + 2),
      borders,
      borderStyles,
      ...(vertAlign ? { vertAlign } : {}),
    })
  }
  return { centers, cells }
}

function vertAlignOf(value: number | undefined): CellVerticalAlign | undefined {
  return value === 1 ? 'center' : value === 2 ? 'bottom' : undefined
}

/** apply `grfbrc`-selected sides of a padding sprm (itcFirst, itcLim, grfbrc, ftsWidth, wWidth) */
function paddingFrom(operand: Uint8Array, padding: TableCellPadding): boolean {
  const grfbrc = operand[2]!
  const ftsWidth = operand[3]!
  const width = u16At(operand, 4)
  // only dxa widths are modelled (ftsWidth nil/auto/percent are not used for margins)
  if (ftsWidth !== 3) return false
  if (grfbrc & 0x01) padding.top = width
  if (grfbrc & 0x02) padding.left = width
  if (grfbrc & 0x04) padding.bottom = width
  if (grfbrc & 0x08) padding.right = width
  return true
}

/** cells itcFirst..itcLim of the row (the TDefTable must already be parsed) */
function cellRange(properties: TapProperties, operand: Uint8Array): TapTableCell[] {
  if (!properties.table || operand.length < 2) return []
  return properties.table.cells.slice(operand[0]!, operand[1]!)
}

function parseTap(grpprl: Uint8Array, rowIndex: number, warnings: string[]): TapProperties {
  const properties: TapProperties = {
    cantSplit: false,
    header: false,
    borders: emptyBorders(),
    borderStyles: {},
  }
  const read = iterateSprmsLenient(grpprl)
  if (read.truncated !== undefined)
    warnings.push(
      'the table properties of row ' +
        rowIndex +
        ' are truncated (' +
        read.truncated +
        '); the properties after that point were ignored',
    )
  for (const sprm of read.sprms) {
    const operation = sprm.opcode & 0x1ff
    switch (operation) {
      case SPRM_T_JC: {
        const value = readSprmU16(sprm.operand)
        if (value === 0) properties.alignment = 'left'
        else if (value === 1) properties.alignment = 'center'
        else if (value === 2) properties.alignment = 'right'
        break
      }
      case SPRM_T_CANT_SPLIT:
        properties.cantSplit = toggleOperand(readSprmU8(sprm.operand), properties.cantSplit)
        break
      case SPRM_T_TABLE_HEADER:
        properties.header = toggleOperand(readSprmU8(sprm.operand), properties.header)
        break
      case SPRM_T_TABLE_BORDERS:
      case SPRM_T_TABLE_BORDERS_WIDE: {
        const wide = operation === SPRM_T_TABLE_BORDERS_WIDE
        if (sprm.operand.length < (wide ? 48 : 24)) {
          warnings.push('the table borders of row ' + rowIndex + ' are truncated and were ignored')
          break
        }
        const set = tableBorderSet(sprm.operand, wide)
        properties.borders = set.borders
        properties.borderStyles = set.styles
        break
      }
      case SPRM_T_SET_BRC80:
      case SPRM_T_SET_BRC: {
        // per-cell border override: itcFirst, itcLim, grfbrc (bit 0 top, 1 left, 2 bottom, 3 right), Brc
        const wide = operation === SPRM_T_SET_BRC
        if (sprm.operand.length < 3 + (wide ? 8 : 4) || !properties.table) break
        const itcFirst = sprm.operand[0]!
        const itcLim = sprm.operand[1]!
        const grfbrc = sprm.operand[2]!
        const line = borderLineAt(sprm.operand, 3, wide)
        for (let itc = itcFirst; itc < itcLim; itc++) {
          const cell = properties.table.cells[itc]
          if (!cell) continue
          ;(['top', 'left', 'bottom', 'right'] as const).forEach((side, sideIndex) => {
            if (!(grfbrc & (1 << sideIndex))) return
            cell.borders[side] = line !== undefined
            if (line) cell.borderStyles[side] = line
            else delete cell.borderStyles[side]
          })
        }
        break
      }
      case SPRM_T_DYA_ROW_HEIGHT: {
        const value = readSprmI16(sprm.operand)
        if (value !== undefined) {
          properties.heightTwips = Math.abs(value)
          properties.heightExact = value < 0
        }
        break
      }
      case SPRM_T_TDEF_TABLE:
        properties.table = parseTDefTable(sprm.operand, rowIndex)
        break
      case SPRM_T_DXA_GAP_HALF: {
        const value = readSprmI16(sprm.operand)
        if (value !== undefined && value >= 0) properties.gapHalf = value
        break
      }
      case SPRM_T_CELL_PADDING_DEFAULT: {
        if (sprm.operand.length < 6) {
          warnings.push('the cell padding of row ' + rowIndex + ' is truncated and was ignored')
          break
        }
        const padding = properties.cellPadding ?? {}
        if (paddingFrom(sprm.operand, padding)) properties.cellPadding = padding
        break
      }
      case SPRM_T_CELL_PADDING: {
        if (sprm.operand.length < 6) break
        for (const cell of cellRange(properties, sprm.operand)) {
          const padding = cell.padding ?? {}
          if (paddingFrom(sprm.operand, padding)) cell.padding = padding
        }
        break
      }
      case SPRM_T_PC: {
        const value = readSprmU8(sprm.operand)
        if (value === undefined) break
        const position = (properties.position ??= {})
        position.vAnchor = POSITION_V_ANCHORS[(value >> 4) & 3] ?? 'margin'
        position.hAnchor = POSITION_H_ANCHORS[(value >> 6) & 3] ?? 'text'
        break
      }
      case SPRM_T_DXA_ABS:
      case SPRM_T_DYA_ABS:
      case SPRM_T_DXA_FROM_TEXT:
      case SPRM_T_DYA_FROM_TEXT:
      case SPRM_T_DXA_FROM_TEXT_RIGHT:
      case SPRM_T_DYA_FROM_TEXT_BOTTOM: {
        const value = readSprmI16(sprm.operand)
        if (value === undefined) break
        const position = (properties.position ??= {})
        if (operation === SPRM_T_DXA_ABS) position.xTwips = value
        else if (operation === SPRM_T_DYA_ABS) position.yTwips = value
        else if (operation === SPRM_T_DXA_FROM_TEXT) position.leftFromText = value
        else if (operation === SPRM_T_DYA_FROM_TEXT) position.topFromText = value
        else if (operation === SPRM_T_DXA_FROM_TEXT_RIGHT) position.rightFromText = value
        else position.bottomFromText = value
        break
      }
      case SPRM_T_F_NO_ALLOW_OVERLAP:
        ;(properties.position ??= {}).noOverlap = toggleOperand(readSprmU8(sprm.operand), false)
        break
      case SPRM_T_VERT_ALIGN: {
        if (sprm.operand.length < 3) break
        const vertAlign = vertAlignOf(sprm.operand[2])
        for (const cell of cellRange(properties, sprm.operand)) {
          if (vertAlign) cell.vertAlign = vertAlign
          else delete cell.vertAlign
        }
        break
      }
      case SPRM_T_DEF_TABLE_SHD80: {
        const cells = properties.table?.cells ?? []
        for (let itc = 0; itc * 2 + 2 <= sprm.operand.length && itc < cells.length; itc++)
          setCellShading(cells[itc]!, shading80At(sprm.operand, itc * 2), false)
        break
      }
      case SPRM_T_SET_SHD:
      case SPRM_T_SET_SHD_ODD: {
        if (sprm.operand.length < 12) break
        const shading = shadingAt(sprm.operand, 2)
        cellRange(properties, sprm.operand).forEach((cell, index) => {
          if (operation === SPRM_T_SET_SHD_ODD && index % 2 === 0) return
          setCellShading(cell, shading, true)
        })
        break
      }
      case SPRM_T_SET_SHD_TABLE: {
        if (sprm.operand.length < 10) break
        properties.shading = shadingAt(sprm.operand, 0)
        break
      }
      default: {
        const first = SPRM_T_DEF_TABLE_SHD[operation] ?? SPRM_T_DEF_TABLE_SHD_RAW[operation]
        if (first !== undefined) {
          const resolved = operation in SPRM_T_DEF_TABLE_SHD
          const cells = properties.table?.cells ?? []
          for (let index = 0; index * 10 + 10 <= sprm.operand.length; index++) {
            const cell = cells[first + index]
            if (cell) setCellShading(cell, shadingAt(sprm.operand, index * 10), resolved)
          }
        }
        // other unknown TAP sprms are skipped by the operand-size rule
        break
      }
    }
  }
  return properties
}

/**
 * Word writes the cell shading three times (Shd80, Shd, raw Shd); the
 * resolved Shd wins and the legacy twins only fill in when it is absent.
 */
function setCellShading(cell: TapTableCell, shading: Shading | undefined, resolved: boolean): void {
  if (cell.shadingResolved && !resolved) return
  if (resolved) cell.shadingResolved = true
  if (shading) cell.shading = shading
  else delete cell.shading
}

/** split the body text (or one story's CP range) into paragraph/cell/row-end units, keeping CP ranges */
function bodyUnits(rawText: string, range?: CpRange): BodyUnit[] {
  const units: BodyUnit[] = []
  const end = range ? Math.min(range.endCp, rawText.length) : rawText.length
  let start = range?.startCp ?? 0
  for (let index = start; index < end; index++) {
    const char = rawText[index]
    if (char === '\r') {
      units.push({
        kind: 'paragraph',
        startCp: start,
        endCp: index,
        text: rawText.slice(start, index),
      })
      start = index + 1
    } else if (char === '\u0007') {
      units.push({ kind: 'cell', startCp: start, endCp: index, text: rawText.slice(start, index) })
      start = index + 1
    }
  }
  // a story range ends on its last paragraph mark; the body keeps its trailing unit
  if (range && start >= end) return units
  units.push({
    kind: 'paragraph',
    startCp: start,
    endCp: end,
    text: rawText.slice(start, end),
  })
  return units
}

export interface CpRange {
  startCp: number
  endCp: number
}

export interface UnitPapx {
  istd: number
  sprms: Uint8Array
}

/** PAPX entries of the whole text keyed by CP (shared by the body and every story built from the same read) */
export type UnitPapxIndex = Map<number, UnitPapx>

export function collectUnitPapx(source: FormattingSource, warnings: string[]): UnitPapxIndex {
  const byCp = new Map<number, UnitPapx>()
  const context: ParagraphContext = { source, warnings }
  const plc = parsePlcBte(source.table, source.fib.fcPlcfbtePapx, source.fib.lcbPlcfbtePapx)
  for (const page of plc.pages) {
    for (const entry of readPapxFkp(source.wordDocument, page)) {
      const resolved = cpOfFc(source.pieces, entry.fcStart)
      if (resolved === 'outside' || resolved === 'unaligned') continue
      byCp.set(resolved.cp, {
        istd: entry.istd,
        sprms: entry.defaults ? EMPTY : expandHugePapx(entry.grpprl, context),
      })
    }
  }
  return byCp
}

interface TableDepth {
  /** sprmPItap (+ sprmPDtap); 1 for a single-level table */
  itap: number
  innerCell: boolean
  innerRowEnd: boolean
}

/** the table nesting level and inner cell/row-end marks a PAPX declares */
function tableDepthOf(papx: Map<number, UnitPapx>, cp: number): TableDepth {
  const depth: TableDepth = { itap: 1, innerCell: false, innerRowEnd: false }
  const entry = papx.get(cp)
  if (!entry) return depth
  let itap: number | undefined
  let delta = 0
  for (const sprm of iterateSprmsLenient(entry.sprms).sprms) {
    switch (sprm.opcode) {
      case SPRM_P_ITAP:
        itap = readSprmI32(sprm.operand) ?? readSprmU16(sprm.operand)
        break
      case SPRM_P_DTAP:
        delta += readSprmI32(sprm.operand) ?? readSprmI16(sprm.operand) ?? 0
        break
      case SPRM_P_F_INNER_TABLE_CELL:
        depth.innerCell = (readSprmU8(sprm.operand) ?? 0) !== 0
        break
      case SPRM_P_F_INNER_TTP:
        depth.innerRowEnd = (readSprmU8(sprm.operand) ?? 0) !== 0
        break
    }
  }
  depth.itap = Math.max((itap ?? 1) + delta, 0)
  return depth
}

function readSprmI32(operand: Uint8Array): number | undefined {
  if (operand.length < 4) return undefined
  return operand[0]! | (operand[1]! << 8) | (operand[2]! << 16) | (operand[3]! << 24)
}

/** paragraph-formatting keys that are structure, not explicit formatting */
interface RawCell {
  units: BodyUnit[]
  definition: TapTableCell
}

interface RawRow {
  rowEnd: BodyUnit
  cells: RawCell[]
  tap: TapProperties
  /** every body unit of the row, for the plain-paragraph fallback */
  units: BodyUnit[]
}

interface RawTable {
  rows: RawRow[]
}

/** attach the row-end TAP to the collected cells; throws TableProblem on a malformed row */
function rawRowFor(
  rowEnd: BodyUnit | undefined,
  cells: RawCell[],
  units: BodyUnit[],
  rowIndex: number,
  papx: Map<number, UnitPapx>,
  warnings: string[],
): RawRow {
  if (!rowEnd) badTable('a table row has no row-end mark')
  if (cells.length === 0) badTable('a table row has no cells')
  const entry = papx.get(rowEnd.startCp)
  if (!entry || entry.sprms.length === 0) badTable('a table row end has no TAP properties')
  const tap = parseTap(entry.sprms, rowIndex, warnings)
  const definition = tap.table
  if (!definition) badTable('the row end of row ' + rowIndex + ' has no TDefTable')
  if (definition.cells.length !== cells.length)
    badTable(
      'row ' +
        rowIndex +
        ' has ' +
        cells.length +
        ' cells but its TDefTable declares ' +
        definition.cells.length,
    )
  cells.forEach((cell, cellIndex) => {
    cell.definition = definition.cells[cellIndex]!
  })
  return { rowEnd, cells, tap, units }
}

type CellSegment = { paragraph: BodyUnit } | { rows: RawRow[]; units: BodyUnit[] }

/**
 * Split a cell's units into its own paragraphs and the inner tables (itap >
 * depth). Inner cells end at a paragraph mark with sprmPFInnerTableCell and
 * inner rows at one with sprmPFInnerTtp, whose PAPX carries the inner TAP.
 */
function splitCellUnits(
  units: BodyUnit[],
  depth: number,
  papx: Map<number, UnitPapx>,
  warnings: string[],
): CellSegment[] {
  const segments: CellSegment[] = []
  let index = 0
  while (index < units.length) {
    if (tableDepthOf(papx, units[index]!.startCp).itap <= depth) {
      segments.push({ paragraph: units[index]! })
      index++
      continue
    }
    const start = index
    const rows: RawRow[] = []
    let cells: RawCell[] = []
    let pending: BodyUnit[] = []
    let rowStart = index
    try {
      while (index < units.length) {
        const unit = units[index]!
        const unitDepth = tableDepthOf(papx, unit.startCp)
        if (unitDepth.itap <= depth) break
        index++
        if (unitDepth.itap !== depth + 1) {
          pending.push(unit)
          continue
        }
        if (unitDepth.innerRowEnd) {
          if (pending.length > 0) cells.push({ units: pending, definition: defaultTapCell() })
          rows.push(
            rawRowFor(unit, cells, units.slice(rowStart, index), rows.length, papx, warnings),
          )
          cells = []
          pending = []
          rowStart = index
        } else if (unitDepth.innerCell) {
          cells.push({ units: [...pending, unit], definition: defaultTapCell() })
          pending = []
        } else pending.push(unit)
      }
      if (cells.length > 0 || pending.length > 0) badTable('a nested table row has no row-end mark')
      if (rows.length === 0) badTable('a nested table has no rows')
      segments.push({ rows, units: units.slice(start, index) })
    } catch (error) {
      if (!(error instanceof TableProblem)) throw error
      warnings.push(
        'a nested table could not be reconstructed (' +
          error.message +
          '); its text was imported as cell paragraphs',
      )
      for (const unit of units.slice(start, index)) segments.push({ paragraph: unit })
    }
  }
  return segments
}

function buildCellParagraph(
  unit: BodyUnit,
  source: FormattingSource,
  read: FormattingRead,
  styles: StyleSheet | undefined,
  papx: Map<number, UnitPapx>,
  warnings: string[],
  fields: SimpleField[] = [],
  noteMarks: NoteMark[] = [],
  commentMarks: CommentMarkAt[] = [],
  pictureCps: number[] = [],
): FormattedParagraph {
  const entry = papx.get(unit.startCp)
  const request: ParagraphRunRequest = {
    startCp: unit.startCp,
    endCp: unit.endCp,
    styleIstd: entry?.istd ?? 0,
    papxSprms: entry?.sprms ?? EMPTY,
  }
  const runs = resolveRunsForRange(
    read,
    source,
    styles,
    request,
    warnings,
    [],
    fieldRangesWithin(fields, unit.startCp, unit.endCp),
    noteMarksWithin(noteMarks, unit.startCp, unit.endCp),
    commentMarksWithin(commentMarks, unit.startCp, unit.endCp),
    pictureCps,
  )
  const text = mapWordControlChars(source.rawText.slice(unit.startCp, unit.endCp))
  const formatting = resolveParagraphFormattingFor(source, styles, request, warnings)
  return {
    startCp: unit.startCp,
    endCp: unit.endCp,
    text,
    formatting,
    runs,
    ...markFormattingEntry(resolveParagraphMarkFormatting(read, source, styles, request, warnings)),
  }
}

/** clearance or no-overlap flags without a position (sprmTPc / TDxaAbs / TDyaAbs) do not float the table */
function isPositionedTable(position: TablePosition | undefined): position is TablePosition {
  return (
    position !== undefined &&
    (position.xTwips !== undefined ||
      position.yTwips !== undefined ||
      position.hAnchor !== undefined)
  )
}

function positionFromFrame(frame: ParagraphFrame | undefined): TablePosition | undefined {
  if (!isPositionedFrame(frame)) return undefined
  return {
    ...(frame.xTwips !== undefined ? { xTwips: frame.xTwips } : {}),
    ...(frame.yTwips !== undefined ? { yTwips: frame.yTwips } : {}),
    ...(frame.hAnchor ? { hAnchor: frame.hAnchor } : {}),
    ...(frame.vAnchor ? { vAnchor: frame.vAnchor } : {}),
    ...(frame.hSpaceTwips
      ? { leftFromText: frame.hSpaceTwips, rightFromText: frame.hSpaceTwips }
      : {}),
    ...(frame.vSpaceTwips
      ? { topFromText: frame.vSpaceTwips, bottomFromText: frame.vSpaceTwips }
      : {}),
  }
}

function buildTable(
  raw: RawTable,
  source: FormattingSource,
  read: FormattingRead,
  styles: StyleSheet | undefined,
  papx: Map<number, UnitPapx>,
  listInfoFor: (
    formatting: ParagraphFormatting,
    directSprms?: Uint8Array,
  ) => ParagraphListInfo | undefined = () => undefined,
  picturesFor: (unit: BodyUnit) => ParagraphPictures | undefined = () => undefined,
  fields: SimpleField[] = [],
  noteMarks: NoteMark[] = [],
  commentMarks: CommentMarkAt[] = [],
  pictureCps: number[] = [],
  depth = 1,
): TableModel {
  const tableWarnings: string[] = []
  // normalize each row's boundaries to its first cell boundary (the table's left edge)
  const rowBoundaries = raw.rows.map((row) => {
    const definition = row.tap.table
    if (!definition) badTable('a table row has no TDefTable')
    const base = definition.centers[0]!
    return definition.centers.map((center) => center - base)
  })
  const gridPoints = [...new Set(rowBoundaries.flat())].sort((left, right) => left - right)
  if (gridPoints.length < 2) badTable('the table has no grid columns')
  const gridWidths: number[] = []
  for (let index = 0; index + 1 < gridPoints.length; index++) {
    const width = gridPoints[index + 1]! - gridPoints[index]!
    if (width <= 0) badTable('the table has an impossible column width')
    gridWidths.push(width)
  }
  const gridIndex = new Map<number, number>()
  gridPoints.forEach((point, index) => gridIndex.set(point, index))
  const rows: TableRow[] = []
  const tableBorders = emptyBorders()
  raw.rows.forEach((row, rowIndex) => {
    const definition = row.tap.table!
    if (row.cells.length !== definition.cells.length)
      badTable(
        'row ' +
          rowIndex +
          ' has ' +
          row.cells.length +
          ' cells but its TDefTable declares ' +
          definition.cells.length,
      )
    const boundaries = rowBoundaries[rowIndex]!
    const cells: TableCell[] = []
    let index = 0
    while (index < row.cells.length) {
      const rawCell = row.cells[index]!
      const startIndex = gridIndex.get(boundaries[index])
      if (startIndex === undefined) badTable('a cell boundary is not on the table grid')
      let endIndex = gridIndex.get(boundaries[index + 1])
      if (endIndex === undefined) badTable('a cell boundary is not on the table grid')
      let last = index
      // fMerged cells are continuations of the previous cell: fold their grid span in
      while (
        last + 1 < row.cells.length &&
        (row.cells[last + 1]!.definition.rgf & TC_MERGED) !== 0
      ) {
        last++
        const mergedEnd = gridIndex.get(boundaries[last + 1])
        if (mergedEnd === undefined) badTable('a merged cell boundary is not on the table grid')
        endIndex = mergedEnd
      }
      const rgf = rawCell.definition.rgf
      const gridSpan = endIndex - startIndex
      if (gridSpan < 1) badTable('a table cell covers no grid column')
      const widthTwips = gridWidths
        .slice(startIndex, endIndex)
        .reduce((sum, width) => sum + width, 0)
      const borders = rawCell.definition.borders
      const cellStyles = rawCell.definition.borderStyles
      if (borders.top || borders.left || borders.bottom || borders.right) {
        tableBorders.top ||= borders.top
        tableBorders.left ||= borders.left
        tableBorders.bottom ||= borders.bottom
        tableBorders.right ||= borders.right
        tableBorders.insideHorizontal ||= borders.top || borders.bottom
        tableBorders.insideVertical ||= borders.left || borders.right
      }
      const paragraphs: FormattedParagraph[] = []
      const paragraphUnits: BodyUnit[] = []
      const nestedTables = new Map<number, TableModel[]>()
      const addParagraph = (unit: BodyUnit): void => {
        paragraphUnits.push(unit)
        paragraphs.push(
          buildCellParagraph(
            unit,
            source,
            read,
            styles,
            papx,
            tableWarnings,
            fields,
            noteMarks,
            commentMarks,
            pictureCps,
          ),
        )
      }
      for (const segment of splitCellUnits(rawCell.units, depth, papx, tableWarnings)) {
        if ('paragraph' in segment) {
          addParagraph(segment.paragraph)
          continue
        }
        try {
          const inner = buildTable(
            { rows: segment.rows },
            source,
            read,
            styles,
            papx,
            listInfoFor,
            picturesFor,
            fields,
            noteMarks,
            commentMarks,
            pictureCps,
            depth + 1,
          )
          tableWarnings.push(...inner.warnings)
          const at = paragraphs.length
          nestedTables.set(at, [...(nestedTables.get(at) ?? []), inner])
        } catch (error) {
          if (!(error instanceof TableProblem)) throw error
          tableWarnings.push(
            'a nested table could not be reconstructed (' +
              error.message +
              '); its text was imported as cell paragraphs',
          )
          for (const unit of segment.units) addParagraph(unit)
        }
      }
      const paragraphLists = new Map<number, ParagraphListInfo>()
      const paragraphPictures = new Map<number, ParagraphPictures>()
      paragraphs.forEach((paragraph, paragraphIndex) => {
        const unit = paragraphUnits[paragraphIndex]!
        const list = listInfoFor(paragraph.formatting, papx.get(unit.startCp)?.sprms ?? EMPTY)
        if (list) paragraphLists.set(paragraphIndex, list)
        const pictures = picturesFor(unit)
        if (pictures) paragraphPictures.set(paragraphIndex, pictures)
      })
      const definitionCell = rawCell.definition
      cells.push({
        gridSpan,
        widthTwips,
        ...(definitionCell.vertAlign ? { vertAlign: definitionCell.vertAlign } : {}),
        ...(definitionCell.shading ? { shading: definitionCell.shading } : {}),
        ...(definitionCell.padding ? { padding: definitionCell.padding } : {}),
        ...(rgf & TC_VERT_RESTART
          ? { verticalMerge: 'restart' as const }
          : rgf & TC_VERT_MERGE
            ? { verticalMerge: 'continue' as const }
            : {}),
        ...(borders.top || borders.left || borders.bottom || borders.right
          ? { borders, ...(Object.keys(cellStyles).length > 0 ? { borderStyles: cellStyles } : {}) }
          : {}),
        paragraphs,
        ...(paragraphLists.size > 0 ? { paragraphLists } : {}),
        ...(paragraphPictures.size > 0 ? { paragraphPictures } : {}),
        ...(nestedTables.size > 0 ? { nestedTables } : {}),
      })
      index = last + 1
    }
    rows.push({
      cells,
      header: row.tap.header,
      cantSplit: row.tap.cantSplit,
      ...(row.tap.heightTwips !== undefined
        ? { heightTwips: row.tap.heightTwips, heightExact: row.tap.heightExact === true }
        : {}),
    })
    tableBorders.top ||= row.tap.borders.top
    tableBorders.left ||= row.tap.borders.left
    tableBorders.bottom ||= row.tap.borders.bottom
    tableBorders.right ||= row.tap.borders.right
    tableBorders.insideHorizontal ||= row.tap.borders.insideHorizontal
    tableBorders.insideVertical ||= row.tap.borders.insideVertical
  })
  const alignment = raw.rows.find((row) => row.tap.alignment !== undefined)?.tap.alignment
  const cellParagraphs = rows.flatMap((row) => row.cells.flatMap((cell) => cell.paragraphs))
  // Word 97 positions a framed table through its paragraphs; the TAP copy wins when both exist
  const position =
    raw.rows.find((row) => isPositionedTable(row.tap.position))?.tap.position ??
    positionFromFrame(
      cellParagraphs.find((paragraph) => isPositionedFrame(paragraph.formatting.frame))?.formatting
        .frame,
    )
  if (position) for (const paragraph of cellParagraphs) delete paragraph.formatting.frame
  const gapHalf = raw.rows.find((row) => row.tap.gapHalf !== undefined)?.tap.gapHalf
  const cellPadding: TableCellPadding = {
    left: gapHalf ?? DEFAULT_GAP_HALF,
    right: gapHalf ?? DEFAULT_GAP_HALF,
    ...raw.rows.find((row) => row.tap.cellPadding !== undefined)?.tap.cellPadding,
  }
  const shading = raw.rows.find((row) => row.tap.shading !== undefined)?.tap.shading
  const borderStyles = raw.rows.find((row) => Object.keys(row.tap.borderStyles).length > 0)?.tap
    .borderStyles
  return {
    gridWidths,
    rows,
    ...(alignment ? { alignment } : {}),
    borders: tableBorders,
    ...(borderStyles ? { borderStyles } : {}),
    cellPadding,
    ...(shading ? { shading } : {}),
    ...(position ? { position } : {}),
    warnings: [...new Set(tableWarnings)],
  }
}

/** visit every cell of a table, including the cells of its nested tables */
export function forEachTableCell(table: TableModel, visit: (cell: TableCell) => void): void {
  for (const row of table.rows)
    for (const cell of row.cells) {
      visit(cell)
      for (const nested of cell.nestedTables?.values() ?? [])
        for (const inner of nested) forEachTableCell(inner, visit)
    }
}

/**
 * Build the reconstructed body (paragraphs and tables in CP order).
 * `read` is the stage-2 formatting read of the same source (optional; it is
 * read here when absent). `styles` resolves character formatting; without it
 * only direct formatting applies.
 */
export function buildLegacyDocBody(
  source: FormattingSource,
  read?: FormattingRead,
  options: {
    styles?: StyleSheet
    images?: ImageBlock[]
    lists?: ListTables
    fields?: FieldParseResult
    /** floating pictures of the body story (drawings.ts) */
    drawings?: FloatingShape[]
    /** footnote/endnote reference marks in the body (notes.ts) */
    noteMarks?: NoteMark[]
    comments?: LegacyComment[]
    /** build only this CP range (a header/footer story) instead of the whole text */
    range?: CpRange
    papx?: UnitPapxIndex
  } = {},
): LegacyDocBody {
  const warnings: string[] = []
  const detailed = read ?? readFormattingDetailed(source, warnings)
  const units = bodyUnits(source.rawText, options.range)
  const papx = options.papx ?? collectUnitPapx(source, warnings)
  const context: ParagraphContext = { source, warnings }
  // classify the \x07 markers: an empty marker after a cell is a row end;
  // paragraph units whose PAPX says inTable belong to a cell (multi-paragraph cell).
  // Documents without cell marks have no table structure to reconstruct even if
  // some PAPX still carries the inTable bit.
  const hasTableMarks = units.some((unit) => unit.kind === 'cell')
  const markerKinds = new Map<number, 'cell' | 'rowEnd'>()
  const cellParagraph = new Set<number>()
  let previousMarker: 'cell' | 'rowEnd' | undefined
  let previousMarkerEmpty = false
  units.forEach((unit, index) => {
    if (unit.kind === 'paragraph') {
      const entry = papx.get(unit.startCp)
      if (entry) {
        const formatting: ParagraphFormatting = {}
        applyParagraphGrpprl(formatting, entry.sprms, context, 'inherit')
        if (formatting.inTable && hasTableMarks) cellParagraph.add(index)
      }
      return
    }
    let kind: 'cell' | 'rowEnd'
    if (unit.text !== '') kind = 'cell'
    else {
      const entry = papx.get(unit.startCp)
      let classified: 'cell' | 'rowEnd' | undefined
      if (entry) {
        const formatting: ParagraphFormatting = {}
        applyParagraphGrpprl(formatting, entry.sprms, context, 'inherit')
        if (formatting.tableRowEnd) classified = 'rowEnd'
        else if (formatting.inTable) classified = 'cell'
      }
      kind = classified ?? (previousMarker === 'cell' && previousMarkerEmpty ? 'rowEnd' : 'cell')
    }
    markerKinds.set(index, kind)
    previousMarker = kind
    previousMarkerEmpty = unit.text === ''
  })
  const blocks: BodyBlock[] = []
  const imagesByCp = new Map<number, ImageBlock>()
  for (const image of options.images ?? []) imagesByCp.set(image.cp, image)
  const pictureCps = [...imagesByCp.keys()]
  const consumedImageCps = new Set<number>()
  const consumedDrawings = new Set<FloatingShape>()
  const drawings = options.drawings ?? []
  const picturesFor = (unit: BodyUnit): ParagraphPictures | undefined => {
    const pictures: ParagraphPictures = {}
    const floating = drawings.filter(
      (drawing) => drawing.cp >= unit.startCp && drawing.cp <= unit.endCp,
    )
    if (floating.length > 0) {
      pictures.floating = floating
      for (const drawing of floating) consumedDrawings.add(drawing)
    }
    const inline: NonNullable<ParagraphPictures['inline']> = []
    for (let cp = unit.startCp; cp < unit.endCp; cp++) {
      const image = imagesByCp.get(cp)
      if (!image || consumedImageCps.has(cp)) continue
      consumedImageCps.add(cp)
      const before = mapWordControlChars(source.rawText.slice(unit.startCp, cp))
      inline.push({ image, position: before === '' ? 'start' : 'end' })
    }
    if (inline.length > 0) pictures.inline = inline
    return pictures.floating || pictures.inline ? pictures : undefined
  }
  const hyperlinks = options.fields?.hyperlinks ?? []
  const linkRangesFor = (startCp: number, endCp: number): FormattedLinkRange[] =>
    hyperlinks
      .filter((link) => link.endCp > startCp && link.startCp < endCp)
      .map((link) => ({
        startCp: Math.max(link.startCp, startCp),
        endCp: Math.min(link.endCp, endCp),
        href: link.anchor ? (link.target ?? '') + '#' + link.anchor : (link.target ?? ''),
        ...(link.tooltip ? { tooltip: link.tooltip } : {}),
        ...(link.inToc ? { inToc: true } : {}),
      }))
      .filter((range) => range.href !== '') as FormattedLinkRange[]
  const simpleFields = options.fields?.simpleFields ?? []
  const noteMarks = options.noteMarks ?? []
  // the reference run follows the \x05 annotation character; a range closes at its bookmark end
  const commentMarks: CommentMarkAt[] = (options.comments ?? []).flatMap((comment) => [
    ...(comment.rangeStartCp !== undefined && comment.rangeEndCp !== undefined
      ? [
          { kind: 'start' as const, id: comment.id, cp: comment.rangeStartCp },
          { kind: 'end' as const, id: comment.id, cp: comment.rangeEndCp },
        ]
      : []),
    {
      kind: 'reference' as const,
      id: comment.id,
      cp: Math.max(comment.referenceCp + 1, comment.rangeEndCp ?? 0),
    },
  ])
  const missingListIlfos = new Set<number>()
  // Word layers the LVL indents between the style and the direct PAPX, so a
  // style indent must not outrank the level's the way a w:ind attribute would
  const applyLevelIndents = (formatting: ParagraphFormatting, directSprms: Uint8Array): void => {
    const ilfo = formatting.listId
    if (ilfo === undefined || ilfo === 0 || !options.lists) return
    const level = resolveListLevel(
      options.lists,
      ilfo,
      Math.min(Math.max(formatting.listLevel ?? 0, 0), 8),
    )
    if (!level) return
    const direct: ParagraphFormatting = {}
    applyParagraphGrpprl(direct, directSprms, context, 'inherit')
    for (const key of ['indentLeftTwips', 'indentFirstLineTwips', 'indentRightTwips'] as const) {
      if (level[key] !== undefined && direct[key] === undefined) formatting[key] = level[key]
    }
  }
  const listInfoFor = (
    formatting: ParagraphFormatting,
    directSprms?: Uint8Array,
  ): ParagraphListInfo | undefined => {
    const ilfo = formatting.listId
    if (ilfo === undefined || ilfo === 0 || !options.lists) return undefined
    if (directSprms) applyLevelIndents(formatting, directSprms)
    const lsid = options.lists.overrides.get(ilfo)
    const definition = lsid !== undefined ? options.lists.lists.get(lsid) : undefined
    if (!definition) {
      if (!missingListIlfos.has(ilfo)) {
        missingListIlfos.add(ilfo)
        warnings.push(
          'a paragraph references a list (ilfo ' +
            ilfo +
            ') that is missing; list formatting was omitted',
        )
      }
      return undefined
    }
    const ilvl = Math.min(Math.max(formatting.listLevel ?? 0, 0), 8)
    const level = resolveListLevel(options.lists, ilfo, ilvl)
    return {
      kind: level && isBulletLevel(level) ? 'bullet' : 'ordered',
      numId: String(100 + ilfo),
      ilvl,
    }
  }
  const paragraphBlockFor = (unit: BodyUnit): BodyBlock => {
    const entry = papx.get(unit.startCp)
    const request: ParagraphRunRequest = {
      startCp: unit.startCp,
      endCp: unit.endCp,
      styleIstd: entry?.istd ?? 0,
      papxSprms: entry?.sprms ?? EMPTY,
    }
    const formatting = resolveParagraphFormattingFor(source, options.styles, request, warnings)
    const list = listInfoFor(formatting, request.papxSprms)
    const pictures = picturesFor(unit)
    return {
      type: 'paragraph',
      ...(list ? { list } : {}),
      ...(pictures ? { pictures } : {}),
      paragraph: {
        startCp: unit.startCp,
        endCp: unit.endCp,
        text: mapWordControlChars(source.rawText.slice(unit.startCp, unit.endCp)),
        formatting,
        runs: resolveRunsForRange(
          detailed,
          source,
          options.styles,
          request,
          warnings,
          linkRangesFor(unit.startCp, unit.endCp),
          fieldRangesWithin(simpleFields, unit.startCp, unit.endCp),
          noteMarksWithin(noteMarks, unit.startCp, unit.endCp),
          commentMarksWithin(commentMarks, unit.startCp, unit.endCp),
          pictureCps,
        ),
        ...markFormattingEntry(
          resolveParagraphMarkFormatting(detailed, source, options.styles, request, warnings),
        ),
      },
    }
  }
  const flushRows = (rows: RawRow[]): void => {
    if (rows.length === 0) return
    try {
      const table = buildTable(
        { rows },
        source,
        detailed,
        options.styles,
        papx,
        listInfoFor,
        picturesFor,
        simpleFields,
        noteMarks,
        commentMarks,
        pictureCps,
      )
      blocks.push({ type: 'table', table })
    } catch (error) {
      if (!(error instanceof TableProblem)) throw error
      warnings.push(
        'a table could not be reconstructed (' +
          error.message +
          '); its text was imported as plain paragraphs',
      )
      for (const row of rows) for (const unit of row.units) blocks.push(paragraphBlockFor(unit))
    }
    rows.length = 0
  }
  let index = 0
  while (index < units.length) {
    const unit = units[index]!
    if (unit.kind === 'paragraph' && !cellParagraph.has(index)) {
      const image = imagesByCp.get(unit.startCp)
      if (image && unit.text === '\u0001') {
        const own = paragraphBlockFor(unit)
        blocks.push({
          type: 'image',
          image,
          ...(own.type === 'paragraph'
            ? { paragraph: own.paragraph, ...(own.list ? { list: own.list } : {}) }
            : {}),
        })
        consumedImageCps.add(image.cp)
        index++
        continue
      }
      blocks.push(paragraphBlockFor(unit))
      index++
      continue
    }
    // a table starts at the first cell marker: collect rows until a row end
    // followed by a normal paragraph (or the end of the text)
    const rawRows: RawRow[] = []
    let finished = false
    const unitIsMarker = (at: number): boolean =>
      at < units.length && units[at]!.kind !== 'paragraph'
    const atTableStart = (at: number): boolean =>
      at < units.length && (unitIsMarker(at) || cellParagraph.has(at))
    while (!finished && index < units.length && atTableStart(index)) {
      const rowStart = index
      const cells: RawCell[] = []
      const pending: BodyUnit[] = []
      let rowEnd: BodyUnit | undefined
      while (index < units.length) {
        const current = units[index]!
        if (current.kind === 'paragraph') {
          pending.push(current)
          index++
          continue
        }
        const markerKind = markerKinds.get(index) ?? 'cell'
        index++
        if (markerKind === 'rowEnd') {
          rowEnd = current
          break
        }
        cells.push({ units: [...pending, current], definition: defaultTapCell() })
        pending.length = 0
      }
      const rowUnits = units.slice(rowStart, index)
      try {
        rawRows.push(rawRowFor(rowEnd, cells, rowUnits, rawRows.length, papx, warnings))
      } catch (error) {
        if (!(error instanceof TableProblem)) throw error
        warnings.push(
          'a table row could not be reconstructed (' +
            error.message +
            '); its text was imported as plain paragraphs',
        )
        flushRows(rawRows)
        for (const rowUnit of rowUnits) blocks.push(paragraphBlockFor(rowUnit))
      }
      if (index >= units.length || (!unitIsMarker(index) && !cellParagraph.has(index)))
        finished = true
    }
    flushRows(rawRows)
  }
  // table-level warnings travel with the body so the conversion reports them
  for (const block of blocks) {
    if (block.type === 'table') warnings.push(...block.table.warnings)
  }
  // content this stage still drops, reported once per document
  const rangeText = options.range
    ? source.rawText.slice(options.range.startCp, options.range.endCp)
    : source.rawText
  if (rangeText.includes('\u0001') && options.images === undefined)
    warnings.push('images are not converted yet; image placeholders are omitted')
  const unusedImages = (options.images ?? []).filter((image) => !consumedImageCps.has(image.cp))
  if (unusedImages.length > 0) warnings.push('a picture outside any paragraph was omitted')
  if (drawings.some((drawing) => !consumedDrawings.has(drawing)))
    warnings.push('a floating picture is anchored outside any paragraph and was omitted')
  // a comment whose reference fell outside every paragraph (a row-end mark) has nothing to anchor to
  const referenced = new Set<number>()
  const collect = (paragraphs: FormattedParagraph[]): void => {
    for (const paragraph of paragraphs)
      for (const run of paragraph.runs)
        for (const mark of run.commentMarks ?? [])
          if (mark.kind === 'reference') referenced.add(mark.id)
  }
  for (const block of blocks) {
    if (block.type === 'table') forEachTableCell(block.table, (cell) => collect(cell.paragraphs))
    else if (block.paragraph) collect([block.paragraph])
  }
  const comments = (options.comments ?? []).filter((comment) => referenced.has(comment.id))
  if (comments.length < (options.comments ?? []).length)
    warnings.push(
      `${(options.comments ?? []).length - comments.length} comment(s) are anchored outside any paragraph and were omitted`,
    )
  return {
    blocks,
    warnings: [...new Set(warnings)],
    ...(options.lists ? { lists: options.lists } : {}),
    ...(comments.length > 0 ? { comments } : {}),
  }
}

/** Read the reconstructed body of a .doc (paragraphs, tables and images in CP order). */
export function readLegacyDocBody(
  source: FormattingSource,
  options: { images?: ImageBlock[]; lists?: ListTables; fields?: FieldParseResult } = {},
): LegacyDocBody {
  const detailedWarnings: string[] = []
  const detailed = readFormattingDetailed(source, detailedWarnings)
  const styles = readStyleSheet(source, detailedWarnings)
  const body = buildLegacyDocBody(source, detailed, {
    styles,
    images: options.images,
    lists: options.lists,
    fields: options.fields,
  })
  return {
    blocks: body.blocks,
    warnings: [...new Set([...detailedWarnings, ...body.warnings])],
    ...(body.lists ? { lists: body.lists } : {}),
  }
}
