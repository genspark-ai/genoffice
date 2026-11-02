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
const SPRM_T_CELL_PADDING = 0x34
/** LibreOffice's table/cell border variants: 0xD612 / 0xD62F / 0xD670 */
const SPRM_T_BORDER_VARIANTS = new Set([0x12, 0x70])
/** sprmPItap: table nesting level stored in the paragraph PAPX */
const SPRM_P_ITAP = 0x6649
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

export interface TableCell {
  /** number of grid columns this cell covers */
  gridSpan: number
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

export interface TableModel {
  /** width of each grid column in twips (start at the first cell boundary) */
  gridWidths: number[]
  rows: TableRow[]
  alignment?: 'left' | 'center' | 'right'
  /** visible table borders */
  borders: TableBorders
  /** line styles of the visible borders (sprmTTableBorders); sides without one are single 0.5pt */
  borderStyles?: BorderStyles
  cellPadding?: TableCellPadding
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
  /** inline pictures that share the paragraph with text; written at its start or end */
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
  cellPadding?: TableCellPadding
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
function borderLineAt(bytes: Uint8Array, offset: number, wide: boolean): BorderLine | undefined {
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

/**
 * LibreOffice writes several table/cell border variants whose operands are
 * arrays of BRC-ish structures with slightly different headers. We do not
 * model them; a visible BRC anywhere in such an operand means the table has
 * visible borders, which the writer flattens to a single-line border set.
 */
function operandHasVisibleBorder(bytes: Uint8Array): boolean {
  for (let offset = 0; offset + 4 <= bytes.length; offset += 4)
    if (visibleBorderCode(bytes, offset)) return true
  return false
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
      cells.push({
        rgf: 0,
        wWidth: 0,
        borders: { top: false, left: false, bottom: false, right: false },
        borderStyles: {},
      })
      continue
    }
    const rgf = u16At(operand, base)
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
    cells.push({ rgf, wWidth: u16At(operand, base + 2), borders, borderStyles })
  }
  return { centers, cells }
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
      case SPRM_T_CELL_PADDING: {
        if (sprm.operand.length < 6) {
          warnings.push('the cell padding of row ' + rowIndex + ' is truncated and was ignored')
          break
        }
        const grfbrc = sprm.operand[2]!
        const ftsWidth = sprm.operand[3]!
        const width = u16At(sprm.operand, 4)
        // only dxa widths are modelled; the fixtures use dxa for every side
        if (ftsWidth === 3 && width > 0) {
          const padding = properties.cellPadding ?? {}
          if (grfbrc & 0x01) padding.top = width
          if (grfbrc & 0x02) padding.left = width
          if (grfbrc & 0x04) padding.bottom = width
          if (grfbrc & 0x08) padding.right = width
          properties.cellPadding = padding
        }
        break
      }
      default:
        if (SPRM_T_BORDER_VARIANTS.has(operation) && operandHasVisibleBorder(sprm.operand)) {
          // an unknown border variant: flatten to the single-line border set
          properties.borders = {
            top: true,
            left: true,
            bottom: true,
            right: true,
            insideHorizontal: true,
            insideVertical: true,
          }
        }
        // other unknown TAP sprms are skipped by the operand-size rule
        break
    }
  }
  return properties
}

/** split the body text into paragraph/cell/row-end units, keeping CP ranges */
function bodyUnits(rawText: string): BodyUnit[] {
  const units: BodyUnit[] = []
  let start = 0
  for (let index = 0; index < rawText.length; index++) {
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
  units.push({
    kind: 'paragraph',
    startCp: start,
    endCp: rawText.length,
    text: rawText.slice(start),
  })
  return units
}

interface UnitPapx {
  istd: number
  sprms: Uint8Array
}

function collectUnitPapx(source: FormattingSource, warnings: string[]): Map<number, UnitPapx> {
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

/** the table nesting level a PAPX declares (sprmPItap, 1 for a single-level table) */
function itapOf(papx: Map<number, UnitPapx>, cp: number): number {
  const entry = papx.get(cp)
  if (!entry) return 1
  for (const sprm of iterateSprmsLenient(entry.sprms).sprms)
    if (sprm.opcode === SPRM_P_ITAP) {
      if (sprm.operand.length >= 4)
        return (
          (sprm.operand[0]! |
            (sprm.operand[1]! << 8) |
            (sprm.operand[2]! << 16) |
            (sprm.operand[3]! << 24)) >>>
          0
        )
      const value = readSprmU16(sprm.operand)
      if (value !== undefined) return value
    }
  return 1
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

function buildTable(
  raw: RawTable,
  source: FormattingSource,
  read: FormattingRead,
  styles: StyleSheet | undefined,
  papx: Map<number, UnitPapx>,
  listInfoFor: (formatting: ParagraphFormatting) => ParagraphListInfo | undefined = () => undefined,
  picturesFor: (unit: BodyUnit) => ParagraphPictures | undefined = () => undefined,
  fields: SimpleField[] = [],
  noteMarks: NoteMark[] = [],
  commentMarks: CommentMarkAt[] = [],
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
      const paragraphs = rawCell.units.map((unit) =>
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
        ),
      )
      const paragraphLists = new Map<number, ParagraphListInfo>()
      const paragraphPictures = new Map<number, ParagraphPictures>()
      paragraphs.forEach((paragraph, paragraphIndex) => {
        const list = listInfoFor(paragraph.formatting)
        if (list) paragraphLists.set(paragraphIndex, list)
        const pictures = picturesFor(rawCell.units[paragraphIndex]!)
        if (pictures) paragraphPictures.set(paragraphIndex, pictures)
      })
      cells.push({
        gridSpan,
        widthTwips,
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
  const cellPadding = raw.rows.find((row) => row.tap.cellPadding !== undefined)?.tap.cellPadding
  const borderStyles = raw.rows.find((row) => Object.keys(row.tap.borderStyles).length > 0)?.tap
    .borderStyles
  return {
    gridWidths,
    rows,
    ...(alignment ? { alignment } : {}),
    borders: tableBorders,
    ...(borderStyles ? { borderStyles } : {}),
    ...(cellPadding ? { cellPadding } : {}),
    warnings: [...new Set(tableWarnings)],
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
  } = {},
): LegacyDocBody {
  const warnings: string[] = []
  const detailed = read ?? readFormattingDetailed(source, warnings)
  const units = bodyUnits(source.rawText)
  const papx = collectUnitPapx(source, warnings)
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
  const listInfoFor = (formatting: ParagraphFormatting): ParagraphListInfo | undefined => {
    const ilfo = formatting.listId
    if (ilfo === undefined || ilfo === 0 || !options.lists) return undefined
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
    const list = listInfoFor(formatting)
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
        cells.push({
          units: [...pending, current],
          definition: {
            rgf: 0,
            wWidth: 0,
            borders: { top: false, left: false, bottom: false, right: false },
            borderStyles: {},
          },
        })
        pending.length = 0
      }
      const rowUnits = units.slice(rowStart, index)
      try {
        if (!rowEnd) badTable('a table row has no row-end mark')
        if (cells.length === 0) badTable('a table row has no cells')
        const entry = papx.get(rowEnd.startCp)
        if (!entry || entry.sprms.length === 0) badTable('a table row end has no TAP properties')
        const tap = parseTap(entry.sprms, rawRows.length, warnings)
        const definition = tap.table
        if (!definition) badTable('the row end of row ' + rawRows.length + ' has no TDefTable')
        if (definition.cells.length !== cells.length)
          badTable(
            'row ' +
              rawRows.length +
              ' has ' +
              cells.length +
              ' cells but its TDefTable declares ' +
              definition.cells.length,
          )
        cells.forEach((cell, cellIndex) => {
          cell.definition = definition.cells[cellIndex]!
        })
        rawRows.push({ rowEnd, cells, tap, units: rowUnits })
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
  // nested tables are detected through sprmPItap on any table-related PAPX
  for (const unit of units) {
    if (unit.kind === 'paragraph') continue
    if (itapOf(papx, unit.startCp) > 1) {
      warnings.push('nested tables are not supported yet; the inner table is flattened to text')
      break
    }
  }
  // content this stage still drops, reported once per document
  if (source.rawText.includes('\u0001') && options.images === undefined)
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
    if (block.type === 'table')
      for (const row of block.table.rows) for (const cell of row.cells) collect(cell.paragraphs)
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
