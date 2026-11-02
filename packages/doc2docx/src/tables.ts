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
 * reconstructed. Malformed TAP/TDefTable data raises 'bad_table'.
 */
import { LegacyDocError, isLegacyDocError } from './errors'
import {
  applyParagraphGrpprl,
  cpOfFc,
  readFormattingDetailed,
  type FormattingRead,
  type FormattedParagraph,
  type FormattingSource,
  type ParagraphFormatting,
} from './formatting'
import { readPapxFkp } from './fkp'
import { parsePlcBte } from './plcf'
import {
  resolveParagraphFormattingFor,
  resolveRunsForRange,
  type ParagraphRunRequest,
} from './resolve'
import { iterateSprms, readSprmI16, readSprmU16, readSprmU8, type Sprm } from './sprm'
import type { ImageBlock } from './images'
import { parseStsh, type StyleSheet } from './styles'
import { mapWordControlChars } from './text'

/* TAP sprm operations (opcode & 0x1FF; sgc 5) */
const SPRM_T_JC = 0x00
const SPRM_T_CANT_SPLIT = 0x03
const SPRM_T_TABLE_HEADER = 0x04
const SPRM_T_TABLE_BORDERS = 0x05
const SPRM_T_DYA_ROW_HEIGHT = 0x07
const SPRM_T_TDEF_TABLE = 0x08
const SPRM_T_CELL_PADDING = 0x34
/** LibreOffice's table/cell border variants: 0xD612 / 0xD62F / 0xD670 */
const SPRM_T_BORDER_VARIANTS = new Set([0x12, 0x2f, 0x70])
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
  /** cell content; always at least one (possibly empty) paragraph */
  paragraphs: FormattedParagraph[]
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
  /** visible table borders; any visible border is written as a single-line set */
  borders: TableBorders
  cellPadding?: TableCellPadding
  warnings: string[]
}

export type BodyBlock =
  | { type: 'paragraph'; paragraph: FormattedParagraph }
  | { type: 'table'; table: TableModel }
  | { type: 'image'; image: ImageBlock }

/** The reconstructed body: paragraphs and tables in CP (document) order. */
export interface LegacyDocBody {
  blocks: BodyBlock[]
  warnings: string[]
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
  cellPadding?: TableCellPadding
}

function badTable(detail: string): never {
  throw new LegacyDocError('bad_table', detail)
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
  const required = 1 + (itcMac + 1) * 2 + itcMac * TC_BYTES
  if (operand.length < required)
    badTable(
      'the TDefTable of row ' +
        rowIndex +
        ' needs ' +
        required +
        ' bytes for ' +
        itcMac +
        ' cells but has ' +
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
  const tcStart = 1 + (itcMac + 1) * 2
  for (let index = 0; index < itcMac; index++) {
    const base = tcStart + index * TC_BYTES
    const rgf = u16At(operand, base)
    const borders: CellBorders = {
      top: visibleBorderCode(operand, base + 4),
      left: visibleBorderCode(operand, base + 8),
      bottom: visibleBorderCode(operand, base + 12),
      right: visibleBorderCode(operand, base + 16),
    }
    cells.push({ rgf, wWidth: u16At(operand, base + 2), borders })
  }
  return { centers, cells }
}

function parseTap(grpprl: Uint8Array, rowIndex: number): TapProperties {
  const properties: TapProperties = { cantSplit: false, header: false, borders: emptyBorders() }
  let sprms: Sprm[]
  try {
    sprms = iterateSprms(grpprl)
  } catch (error) {
    if (isLegacyDocError(error))
      badTable('the TAP of row ' + rowIndex + ' is invalid: ' + error.message)
    throw error
  }
  for (const sprm of sprms) {
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
      case SPRM_T_TABLE_BORDERS: {
        if (sprm.operand.length < 24)
          badTable('the table borders of row ' + rowIndex + ' are truncated')
        properties.borders = {
          top: visibleBorderCode(sprm.operand, 0),
          left: visibleBorderCode(sprm.operand, 4),
          bottom: visibleBorderCode(sprm.operand, 8),
          right: visibleBorderCode(sprm.operand, 12),
          insideHorizontal: visibleBorderCode(sprm.operand, 16),
          insideVertical: visibleBorderCode(sprm.operand, 20),
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
        if (sprm.operand.length < 6)
          badTable('the cell padding of row ' + rowIndex + ' is truncated')
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

function collectUnitPapx(source: FormattingSource): Map<number, UnitPapx> {
  const byCp = new Map<number, UnitPapx>()
  const plc = parsePlcBte(source.table, source.fib.fcPlcfbtePapx, source.fib.lcbPlcfbtePapx)
  for (const page of plc.pages) {
    for (const entry of readPapxFkp(source.wordDocument, page)) {
      const resolved = cpOfFc(source.pieces, entry.fcStart)
      if (resolved === 'outside' || resolved === 'unaligned') continue
      byCp.set(resolved.cp, { istd: entry.istd, sprms: entry.defaults ? EMPTY : entry.grpprl })
    }
  }
  return byCp
}

/** the table nesting level a PAPX declares (sprmPItap, 1 for a single-level table) */
function itapOf(papx: Map<number, UnitPapx>, cp: number): number {
  const entry = papx.get(cp)
  if (!entry) return 1
  for (const sprm of iterateSprms(entry.sprms))
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
const STRUCTURAL_PARAGRAPH_KEYS = new Set(['styleIndex', 'inTable', 'tableRowEnd'])

function hasExplicitParagraphFormatting(formatting: ParagraphFormatting): boolean {
  return Object.keys(formatting).some((key) => !STRUCTURAL_PARAGRAPH_KEYS.has(key))
}

interface RawCell {
  units: BodyUnit[]
  definition: TapTableCell
}

interface RawTable {
  rows: Array<{ rowEnd: BodyUnit; cells: RawCell[]; tap: TapProperties }>
}

function buildCellParagraph(
  unit: BodyUnit,
  source: FormattingSource,
  read: FormattingRead,
  styles: StyleSheet | undefined,
  papx: Map<number, UnitPapx>,
  warnings: string[],
): FormattedParagraph {
  const entry = papx.get(unit.startCp)
  const request: ParagraphRunRequest = {
    startCp: unit.startCp,
    endCp: unit.endCp,
    styleIstd: entry?.istd ?? 0,
    papxSprms: entry?.sprms ?? EMPTY,
  }
  const runs = resolveRunsForRange(read, source, styles, request, warnings)
  const text = mapWordControlChars(source.rawText.slice(unit.startCp, unit.endCp))
  // cell paragraph formatting is plain: report explicit PAPX properties once
  const explicit = resolveParagraphFormattingFor(source, styles, request, [])
  if (hasExplicitParagraphFormatting(explicit))
    warnings.push('table cell paragraph formatting is flattened to plain paragraphs')
  return { startCp: unit.startCp, endCp: unit.endCp, text, formatting: {}, runs }
}

function buildTable(
  raw: RawTable,
  source: FormattingSource,
  read: FormattingRead,
  styles: StyleSheet | undefined,
  papx: Map<number, UnitPapx>,
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
      if (borders.top || borders.left || borders.bottom || borders.right) {
        tableBorders.top ||= borders.top
        tableBorders.left ||= borders.left
        tableBorders.bottom ||= borders.bottom
        tableBorders.right ||= borders.right
        tableBorders.insideHorizontal ||= borders.top || borders.bottom
        tableBorders.insideVertical ||= borders.left || borders.right
      }
      const paragraphs = rawCell.units.map((unit) =>
        buildCellParagraph(unit, source, read, styles, papx, tableWarnings),
      )
      cells.push({
        gridSpan,
        widthTwips,
        ...(rgf & TC_VERT_RESTART ? { verticalMerge: 'restart' as const } : {}),
        ...(rgf & TC_VERT_MERGE ? { verticalMerge: 'continue' as const } : {}),
        ...(borders.top || borders.left || borders.bottom || borders.right ? { borders } : {}),
        paragraphs,
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
  return {
    gridWidths,
    rows,
    ...(alignment ? { alignment } : {}),
    borders: tableBorders,
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
  options: { styles?: StyleSheet; images?: ImageBlock[] } = {},
): LegacyDocBody {
  const warnings: string[] = []
  const detailed = read ?? readFormattingDetailed(source, warnings)
  const units = bodyUnits(source.rawText)
  const papx = collectUnitPapx(source)
  const context = { source, warnings }
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
  let index = 0
  while (index < units.length) {
    const unit = units[index]!
    if (unit.kind === 'paragraph' && !cellParagraph.has(index)) {
      const image = imagesByCp.get(unit.startCp)
      if (image && unit.text === '\u0001') {
        blocks.push({ type: 'image', image })
        consumedImageCps.add(image.cp)
        index++
        continue
      }
      const entry = papx.get(unit.startCp)
      const request: ParagraphRunRequest = {
        startCp: unit.startCp,
        endCp: unit.endCp,
        styleIstd: entry?.istd ?? 0,
        papxSprms: entry?.sprms ?? EMPTY,
      }
      blocks.push({
        type: 'paragraph',
        paragraph: {
          startCp: unit.startCp,
          endCp: unit.endCp,
          text: mapWordControlChars(source.rawText.slice(unit.startCp, unit.endCp)),
          formatting: resolveParagraphFormattingFor(source, options.styles, request, warnings),
          runs: resolveRunsForRange(detailed, source, options.styles, request, warnings),
        },
      })
      index++
      continue
    }
    // a table starts at the first cell marker: collect rows until a row end
    // followed by a normal paragraph (or the end of the text)
    const rawRows: RawTable['rows'] = []
    let finished = false
    const unitIsMarker = (at: number): boolean =>
      at < units.length && units[at]!.kind !== 'paragraph'
    const atTableStart = (at: number): boolean =>
      at < units.length && (unitIsMarker(at) || cellParagraph.has(at))
    while (!finished && index < units.length && atTableStart(index)) {
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
          },
        })
        pending.length = 0
      }
      if (!rowEnd) badTable('a table row has no row-end mark')
      if (cells.length === 0) badTable('a table row has no cells')
      const entry = papx.get(rowEnd.startCp)
      if (!entry || entry.sprms.length === 0) badTable('a table row end has no TAP properties')
      const tap = parseTap(entry.sprms, rawRows.length)
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
      rawRows.push({ rowEnd, cells, tap })
      if (index >= units.length || (!unitIsMarker(index) && !cellParagraph.has(index)))
        finished = true
    }
    const table = buildTable({ rows: rawRows }, source, detailed, options.styles, papx)
    blocks.push({ type: 'table', table })
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
  if (source.rawText.includes('\u0001') && (options.images ?? []).length === 0)
    warnings.push('images are not converted yet; image placeholders are omitted')
  if (source.rawText.includes('HYPERLINK'))
    warnings.push(
      'hyperlinks are not converted yet; the link text is imported without the hyperlink',
    )
  const hasListItems = blocks.some(
    (block) =>
      block.type === 'paragraph' &&
      (block.paragraph.formatting.listLevel !== undefined ||
        block.paragraph.formatting.listId !== undefined),
  )
  if (hasListItems)
    warnings.push(
      'numbered and bulleted lists are not converted yet; list text is imported as plain paragraphs',
    )
  const unusedImages = (options.images ?? []).filter((image) => !consumedImageCps.has(image.cp))
  if (unusedImages.length > 0)
    warnings.push(
      'a picture inside a table cell or a text paragraph is not supported yet; it was omitted',
    )
  return { blocks, warnings: [...new Set(warnings)] }
}

/** Read the reconstructed body of a .doc (paragraphs, tables and images in CP order). */
export function readLegacyDocBody(
  source: FormattingSource,
  options: { images?: ImageBlock[] } = {},
): LegacyDocBody {
  const detailedWarnings: string[] = []
  const detailed = readFormattingDetailed(source, detailedWarnings)
  const styles =
    source.fib.lcbStshf > 0
      ? parseStsh(source.table, source.fib.fcStshf, source.fib.lcbStshf)
      : undefined
  const body = buildLegacyDocBody(source, detailed, { styles, images: options.images })
  return {
    blocks: body.blocks,
    warnings: [...new Set([...detailedWarnings, ...body.warnings])],
  }
}
