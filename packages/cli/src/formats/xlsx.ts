import JSZip from 'jszip'
import * as numfmt from 'numfmt'
import {
  blankXlsxBuffer,
  decodeCsvBuffer,
  isNumericCell,
  parseCsv,
} from '@genoffice/xlsx-gateway/gateway/csv-import'
import { columnLabel, parseAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import {
  assembleWithJsZip,
  createBufferEntrySource,
  planCellEditsToXlsx,
  toA1Address,
  type CellEdit,
  type SheetFormulaValues,
  type SheetStructuralOps,
  type XlsxMutation,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import {
  nextFreeRelationshipId,
  type SheetEditPlan,
} from '@genoffice/xlsx-gateway/gateway/xlsx-sheets'
import { EMPTY_PAYLOADS, type GatewayPayloads } from './xlsx-gateway-ops'
import type { WorkbookStyleEdit } from '@genoffice/xlsx-gateway/shared/edit-schemas'
import {
  echoStyleColor,
  fillDisplayColor,
  isGradientFill,
  isThemeColor,
  resolveStyleColor,
  type StyleColor,
  type StyleColorEcho,
} from '@genoffice/xlsx-gateway/domain/style-color'
import {
  parseStylesheetFormats,
  type StylesheetFormats,
} from '@genoffice/xlsx-gateway/gateway/xlsx-style-read'
import { readFile } from 'node:fs/promises'
import { atomicWriteFile } from '../../../../apps/sheets/src/main/atomic-write'
import { XlsxSidecarClient } from '../../../../apps/sheets/src/main/xlsx-sidecar-client'
import { CliError, EXIT } from '../result'
import { sheetNotFoundHints } from '../suggest'
import { xlsxSidecarPath } from '../resources'

export type Scalar = string | number | boolean | null

interface CellArea {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}

/** The subset of the sidecar's workbook file model (apps/sheets desktop-api) the CLI reads. */
interface OpenedWorkbook {
  sessionId: string
  sheets: {
    id: string
    name: string
    rowCount: number
    columnCount: number
    columnWidths?: {
      startColumn: number
      endColumn: number
      width?: number
      hidden: boolean
      styleIndex?: number
    }[]
    freeze?: { frozenRows: number; frozenColumns: number } | null
    hidden?: boolean
    tables?: { range: CellArea; name?: string }[]
    comments?: { row: number; column: number }[]
    pivotTables?: { path: string; outputRef: string }[]
  }[]
  activeTab: number
  definedNames?: unknown[]
  styles?: CellStyle[]
  /** #RRGGBB in theme index order [lt1, dk1, lt2, dk2, accent1-6, hlink, folHlink] */
  themeColors?: string[]
  visuals?: {
    id: string
    sheetId: string
    kind: string
    anchor: { fromRow: number; fromColumn: number }
    chart?: { chartTypes: string[]; title: string; series: unknown[] }
  }[]
  date1904?: boolean
}

interface CellStyle {
  fontFamily?: string
  fontSize?: number
  bold: boolean
  italic: boolean
  underline: boolean
  strikethrough: boolean
  wrapText: boolean
  fontColor?: string
  fillColor?: string
  horizontalAlignment?: string
  verticalAlignment?: string
  numberFormat?: string
}

interface RangeResult {
  cells: { row: number; column: number; value?: Scalar; formula?: string; styleIndex?: number }[]
  rows?: { row: number; styleIndex?: number; height?: number; hidden?: boolean }[]
  merges?: CellArea[]
  hyperlinks?: unknown[]
  conditionalRules?: unknown[]
  autoFilter?: CellArea | null
  dataValidations?: unknown[]
  /** the sidecar indexes worksheets in the background and answers with what is ready so far */
  indexingComplete?: boolean
  indexedThroughRow?: number | null
}

const INDEXING_DEADLINE_MS = 120_000

/**
 * read_range answers after at most 750ms of indexing and expects the caller
 * to poll (the app's merge-workbooks loop); a band read once could come back
 * with rows still missing and no error.
 */
async function readRangeIndexed(
  client: XlsxSidecarClient,
  sessionId: string,
  sheetId: string,
  range: Bounds,
  /** sheet-wide features (filter, rules, links) arrive only with the finished index */
  complete = false,
): Promise<RangeResult> {
  const deadline = Date.now() + INDEXING_DEADLINE_MS
  for (;;) {
    const r = (await client.readRange({ sessionId, sheetId, range })) as RangeResult
    if (
      r.indexingComplete ||
      r.indexedThroughRow === undefined ||
      (!complete && r.indexedThroughRow !== null && r.indexedThroughRow >= range.endRow)
    ) {
      return r
    }
    if (Date.now() > deadline) {
      throw new CliError(EXIT.conversion, 'the workbook did not finish indexing in time', {
        indexed_through_row: r.indexedThroughRow,
        wanted_through_row: range.endRow,
      })
    }
    await new Promise((resolve) => setTimeout(resolve, 400))
  }
}

type SheetMeta = OpenedWorkbook['sheets'][number]

function pickSheet(wb: OpenedWorkbook, name: string | undefined): SheetMeta {
  const meta = name
    ? wb.sheets.find((s) => s.name === name)
    : (wb.sheets[wb.activeTab] ?? wb.sheets[0])
  if (!meta) {
    throw new CliError(
      EXIT.usage,
      `sheet not found: ${name}`,
      { sheets: wb.sheets.map((s) => s.name) },
      sheetNotFoundHints(
        name ?? '',
        wb.sheets.map((s) => s.name),
      ),
    )
  }
  return meta
}

interface RecalcResult {
  cells: {
    sheet: string
    row: number
    column: number
    formatted: string
    number?: number
    isError?: boolean
  }[]
}

/** One short-lived sidecar process per call; the CLI has no session to keep alive. */
async function withSidecar<T>(fn: (client: XlsxSidecarClient) => Promise<T>): Promise<T> {
  const binary = xlsxSidecarPath()
  if (!binary) {
    throw new CliError(EXIT.conversion, 'xlsx engine (xlsx-sidecar) not found', {
      hint: 'set XLSX_SIDECAR_PATH or run a packaged GenOffice',
    })
  }
  const client = new XlsxSidecarClient(binary)
  try {
    return await fn(client)
  } finally {
    client.stop()
  }
}

async function withOpenWorkbook<T>(
  path: string,
  fn: (client: XlsxSidecarClient, wb: OpenedWorkbook) => Promise<T>,
): Promise<T> {
  return withSidecar(async (client) => {
    const wb = (await client.open(path, 'en')) as OpenedWorkbook
    try {
      return await fn(client, wb)
    } finally {
      await client.close(wb.sessionId)
    }
  })
}

export interface SheetSummary {
  name: string
  rows: number
  columns: number
}

export interface WorkbookSummary {
  sheets: SheetSummary[]
  definedNames: number
  activeSheet: string | null
}

export async function workbookSummary(path: string): Promise<WorkbookSummary> {
  return withOpenWorkbook(path, async (_client, wb) => ({
    sheets: wb.sheets.map((s) => ({ name: s.name, rows: s.rowCount, columns: s.columnCount })),
    definedNames: wb.definedNames?.length ?? 0,
    activeSheet: wb.sheets[wb.activeTab]?.name ?? null,
  }))
}

/** Legacy spreadsheet formats (xls, xlsb, ods) → xlsx with values and formulas. */
export async function convertLegacyWorkbook(
  path: string,
  targetPath: string,
): Promise<{ sheets: number; cells: number }> {
  return withSidecar(async (client) => {
    return (await client.convertWorkbook({ path, targetPath })) as { sheets: number; cells: number }
  })
}

// ── reading ────────────────────────────────────────────────────────────

const READ_ROW_CAP = 500
const READ_COL_CAP = 100

export interface SheetFeatures {
  frozen: { rows: number; columns: number } | null
  filter: string | null
  /** merged ranges that touch the read range */
  merges: string[]
  charts: { id: string; title: string; types: string[]; series: number; anchor: string }[]
  tables: { name: string | null; range: string }[]
  /** output ranges of the sheet's pivot tables */
  pivots: string[]
  conditionalFormats: number
  dataValidations: number
  /** links on cells of the read range */
  hyperlinks: number
  notes: number
  hiddenColumns: number
  hiddenSheet: boolean
}

export interface CellFormat {
  bold?: true
  italic?: true
  underline?: true
  strikethrough?: true
  wrapText?: true
  fontFamily?: string
  fontSize?: number
  fontColor?: string
  /** when the font color is a theme slot: the slot and tint behind `fontColor` */
  fontColorTheme?: { theme: string; tint?: number }
  fillColor?: string
  /** pattern / gradient / theme-slot fills as stored, with each color resolved to rgb */
  fill?: FillEcho
  horizontalAlign?: string
  verticalAlign?: string
  numberFormat?: string
}

export type FillEcho =
  | { pattern: string; fg: StyleColorEcho; bg?: StyleColorEcho }
  | {
      gradient: {
        type?: 'linear' | 'path'
        angle?: number
        left?: number
        right?: number
        top?: number
        bottom?: number
        stops: { position: number; color: StyleColorEcho }[]
      }
    }
export type ReadWhere = 'formula' | 'error' | 'empty' | 'number' | 'text'
export const READ_WHERE: readonly ReadWhere[] = ['formula', 'error', 'empty', 'number', 'text']

const ERROR_VALUE = /^#(N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA)$/

export interface ReadOptions {
  sheet?: string
  range?: string
  formats?: boolean
  /** columns to keep, `A,C,E:G` */
  cols?: string
  /** stop after this many rows of the range */
  maxRows?: number
  where?: ReadWhere
  stats?: boolean
}

export interface MatchedCell {
  ref: string
  value: Scalar
  formula?: string
}

export interface SheetStats {
  sheets: { name: string; rows: number; columns: number }[]
  usedRange: string | null
  rows: number
  columns: number
  /** the range the counts below were taken over (the whole used area unless capped or --range) */
  scanned: string
  nonEmpty: number
  formulas: number
  errors: number
  numbers: number
  text: number
  booleans: number
  merges: number
}

export interface SheetRead {
  sheet: string
  range: string
  rows: Scalar[][]
  formulas: Record<string, string>
  truncated: boolean
  rowsShown: number
  rowsTotal: number
  /** with `cols`: the column of each entry of every row */
  columns?: string[]
  /** with `where`: only the matching cells */
  cells?: MatchedCell[]
  matches?: number
  stats?: SheetStats
  features: SheetFeatures
  /** with `formats`: styled cells of the range, by A1 address */
  formats?: Record<string, CellFormat>
  /** with `formats`: column widths in px (Calibri 11 characters × 7) for columns the file sizes */
  columnWidths?: Record<string, number>
  /** with `formats`: row heights in points for rows the file sizes */
  rowHeights?: Record<string, number>
}

/** Excel's sheet limits; the gateway would write anything beyond them into a file no reader accepts. */
const MAX_ROWS = 1_048_576
const MAX_COLUMNS = 16_384

export function parseAddressChecked(
  address: string,
  what = 'cell',
): { row: number; column: number } {
  let parsed: { row: number; column: number }
  try {
    parsed = parseAddress(address.toUpperCase())
  } catch {
    throw new CliError(EXIT.usage, `invalid ${what} address: ${address}`, undefined, {
      reason: 'invalid_argument',
    })
  }
  if (parsed.row >= MAX_ROWS || parsed.column >= MAX_COLUMNS) {
    throw new CliError(
      EXIT.usage,
      `${what} address outside the sheet limits (XFD1048576): ${address}`,
    )
  }
  return parsed
}

const MATCH_CAP = 1000

function parseCols(spec: string): number[] {
  const bad = (): CliError =>
    new CliError(EXIT.usage, `--cols must look like A,C,E:G (got ${spec})`, undefined, {
      reason: 'invalid_argument',
    })
  const out = new Set<number>()
  for (const part of spec.split(',')) {
    const m = /^([A-Za-z]{1,3})(?::([A-Za-z]{1,3}))?$/.exec(part.trim())
    if (!m) throw bad()
    const a = parseAddressChecked(`${m[1]}1`, 'column').column
    const b = m[2] ? parseAddressChecked(`${m[2]}1`, 'column').column : a
    for (let c = Math.min(a, b); c <= Math.max(a, b); c++) out.add(c)
  }
  if (!out.size) throw bad()
  return [...out].sort((x, y) => x - y)
}

function parseRangeChecked(range: string): Bounds {
  const parts = range.split(':')
  if (parts.length > 2 || !parts[0])
    throw new CliError(EXIT.usage, `invalid range: ${range}`, undefined, {
      reason: 'invalid_argument',
    })
  const first = parseAddressChecked(parts[0], 'range')
  const second = parts[1] ? parseAddressChecked(parts[1], 'range') : first
  return {
    startRow: Math.min(first.row, second.row),
    startColumn: Math.min(first.column, second.column),
    endRow: Math.max(first.row, second.row),
    endColumn: Math.max(first.column, second.column),
  }
}

export async function readSheet(path: string, opts: ReadOptions): Promise<SheetRead> {
  const cols = opts.cols ? parseCols(opts.cols) : undefined
  return withOpenWorkbook(path, async (client, wb) => {
    const meta = pickSheet(wb, opts.sheet)
    const bounds = opts.range
      ? parseRangeChecked(opts.range)
      : {
          startRow: 0,
          startColumn: 0,
          endRow: Math.max(0, Math.min(meta.rowCount, READ_ROW_CAP) - 1),
          endColumn: Math.max(0, Math.min(meta.columnCount, READ_COL_CAP) - 1),
        }
    const truncated =
      !opts.range && (meta.rowCount > READ_ROW_CAP || meta.columnCount > READ_COL_CAP)
    // a range beyond the used area is simply empty, not an error; the sheet-wide
    // features still come with the read, so the request is clamped, not skipped
    const empty =
      meta.rowCount === 0 ||
      meta.columnCount === 0 ||
      bounds.startRow >= meta.rowCount ||
      bounds.startColumn >= meta.columnCount
    const result: RangeResult = await readRangeIndexed(
      client,
      wb.sessionId,
      meta.id,
      empty
        ? { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }
        : {
            startRow: bounds.startRow,
            endRow: Math.min(bounds.endRow, Math.max(meta.rowCount - 1, bounds.startRow)),
            startColumn: bounds.startColumn,
            endColumn: Math.min(
              bounds.endColumn,
              Math.max(meta.columnCount - 1, bounds.startColumn),
            ),
          },
      true,
    )
    if (empty) {
      // the A1 probe only serves the sheet-wide fields; nothing per-range belongs to the asked area
      result.cells = []
      result.rows = []
      result.merges = []
      result.hyperlinks = []
    }
    const height = bounds.endRow - bounds.startRow + 1
    const width = bounds.endColumn - bounds.startColumn + 1
    const rows: Scalar[][] = Array.from({ length: height }, () => Array<Scalar>(width).fill(null))
    const formulas: Record<string, string> = {}
    const pending: { row: number; column: number }[] = []
    for (const c of result.cells) {
      const r = c.row - bounds.startRow
      const k = c.column - bounds.startColumn
      if (r < 0 || k < 0 || r >= height || k >= width) continue
      rows[r]![k] = c.value ?? null
      if (c.formula) {
        formulas[`${columnLabel(c.column)}${c.row + 1}`] = c.formula.startsWith('=')
          ? c.formula
          : `=${c.formula}`
        if (c.value === undefined || c.value === null)
          pending.push({ row: c.row, column: c.column })
      }
    }
    // formulas written without a cached value (e.g. by genoffice itself) are evaluated on the fly
    if (pending.length) {
      const box = boundingBox(pending)
      const cells = await recalcRange(client, path, meta.name, box)
      for (const cell of cells) {
        const r = cell.row - bounds.startRow
        const k = cell.column - bounds.startColumn
        if (r < 0 || k < 0 || r >= height || k >= width) continue
        if (rows[r]![k] === null && cell.formatted !== '')
          rows[r]![k] = cell.number ?? cell.formatted
      }
    }
    const rangeLabel = `${columnLabel(bounds.startColumn)}${bounds.startRow + 1}:${columnLabel(bounds.endColumn)}${bounds.endRow + 1}`
    const stats = opts.stats ? sheetStats(wb, meta, result, rows, formulas, rangeLabel) : undefined
    const rowsTotal = opts.range ? height : Math.max(meta.rowCount, height)
    const shownRows =
      opts.maxRows !== undefined && opts.maxRows < rows.length ? rows.slice(0, opts.maxRows) : rows
    const lastRow = bounds.startRow + Math.max(shownRows.length, 1) - 1
    const keptColumns = cols?.filter((c) => c >= bounds.startColumn && c <= bounds.endColumn)
    if (keptColumns && !keptColumns.length) {
      throw new CliError(
        EXIT.usage,
        `--cols ${opts.cols} lies outside ${rangeLabel}`,
        { valid_range: rangeLabel },
        { reason: 'out_of_range', suggestion: 'widen --range or drop --cols' },
      )
    }
    const inScope = (address: string): boolean => {
      const a = parseAddress(address)
      return a.row <= lastRow && (!keptColumns || keptColumns.includes(a.column))
    }
    const pick = <T>(byAddress: Record<string, T>): Record<string, T> =>
      Object.fromEntries(Object.entries(byAddress).filter(([k]) => inScope(k)))
    const read: SheetRead = {
      sheet: meta.name,
      range:
        shownRows.length < rows.length
          ? `${columnLabel(bounds.startColumn)}${bounds.startRow + 1}:${columnLabel(bounds.endColumn)}${lastRow + 1}`
          : rangeLabel,
      rows: keptColumns
        ? shownRows.map((r) => keptColumns.map((c) => r[c - bounds.startColumn] ?? null))
        : shownRows,
      formulas: pick(formulas),
      truncated: truncated || shownRows.length < rowsTotal,
      rowsShown: shownRows.length,
      rowsTotal,
      ...(keptColumns ? { columns: keptColumns.map(columnLabel) } : {}),
      features: sheetFeatures(wb, meta, result),
    }
    if (opts.where) {
      const cells: MatchedCell[] = []
      const columns = keptColumns ?? Array.from({ length: width }, (_, k) => bounds.startColumn + k)
      let matches = 0
      shownRows.forEach((r, i) => {
        columns.forEach((c) => {
          const value = r[c - bounds.startColumn] ?? null
          const ref = `${columnLabel(c)}${bounds.startRow + i + 1}`
          const formula = formulas[ref]
          if (!cellMatches(opts.where!, value, formula)) return
          matches++
          if (cells.length < MATCH_CAP) cells.push({ ref, value, ...(formula ? { formula } : {}) })
        })
      })
      read.cells = cells
      read.matches = matches
    }
    if (stats) read.stats = stats
    if (opts.formats) {
      const formats: Record<string, CellFormat> = {}
      // xf 0 is the workbook default; a style only counts where it differs from it
      const base = wb.styles?.[0]
      const stored = await storedFormats(path)
      for (const c of result.cells) {
        if (!c.styleIndex) continue
        const style = wb.styles?.[c.styleIndex]
        const format = style ? cellFormat(style, base) : undefined
        const echoed = withStoredColors(format, stored, c.styleIndex, wb.themeColors)
        if (echoed) formats[`${columnLabel(c.column)}${c.row + 1}`] = echoed
      }
      read.formats = pick(formats)
      read.columnWidths = Object.fromEntries(
        (meta.columnWidths ?? [])
          .filter((w) => w.width !== undefined)
          .flatMap((w) => {
            const out: [string, number][] = []
            for (let c = w.startColumn; c <= Math.min(w.endColumn, bounds.endColumn); c++) {
              if (c >= bounds.startColumn && (!keptColumns || keptColumns.includes(c)))
                out.push([columnLabel(c), Math.round(w.width! * PX_PER_CHAR)])
            }
            return out
          }),
      )
      read.rowHeights = Object.fromEntries(
        (result.rows ?? [])
          .filter((r) => r.height !== undefined && r.row <= lastRow)
          .map((r) => [String(r.row + 1), r.height!]),
      )
    }
    return read
  })
}

async function hasThemePart(source: Buffer): Promise<boolean> {
  const zip = await JSZip.loadAsync(source)
  return zip.file('xl/theme/theme1.xml') !== null
}

function styleHasThemeColor(style: WorkbookStyleEdit): boolean {
  const colors: unknown[] = [style.fontColor, style.fillColor]
  for (const edge of [style.borderTop, style.borderBottom, style.borderLeft, style.borderRight]) {
    if (edge) colors.push(edge.color)
  }
  if (style.fill) {
    if ('gradient' in style.fill) colors.push(...style.fill.gradient.stops.map((s) => s.color))
    else colors.push(style.fill.fg, style.fill.bg)
  }
  return colors.some((c) => isThemeColor(c as StyleColor | null | undefined))
}

function flattenThemeColors(style: WorkbookStyleEdit): WorkbookStyleEdit {
  const rgb = <T extends StyleColor | null | undefined>(c: T): string | T =>
    isThemeColor(c) ? resolveStyleColor(c) : c
  const out: WorkbookStyleEdit = { ...style }
  if (out.fontColor !== undefined) out.fontColor = rgb(out.fontColor)
  if (out.fillColor !== undefined) out.fillColor = rgb(out.fillColor)
  for (const key of ['borderTop', 'borderBottom', 'borderLeft', 'borderRight'] as const) {
    const edge = out[key]
    if (edge && edge.color !== undefined) out[key] = { ...edge, color: rgb(edge.color) }
  }
  if (out.fill) {
    out.fill =
      'gradient' in out.fill
        ? {
            gradient: {
              ...out.fill.gradient,
              stops: out.fill.gradient.stops.map((s) => ({ ...s, color: rgb(s.color) })),
            },
          }
        : {
            ...out.fill,
            fg: rgb(out.fill.fg),
            ...(out.fill.bg === undefined ? {} : { bg: rgb(out.fill.bg) }),
          }
  }
  return out
}

function sameStyleColor(a: StyleColor | undefined, b: StyleColor | undefined): boolean {
  if (typeof a === 'string' || typeof b === 'string' || !a || !b) return a === b
  return a.theme === b.theme && (a.tint ?? 0) === (b.tint ?? 0)
}

async function storedFormats(path: string): Promise<StylesheetFormats | undefined> {
  try {
    const zip = await JSZip.loadAsync(await readFile(path))
    const xml = await zip.file('xl/styles.xml')?.async('string')
    return xml === undefined ? undefined : parseStylesheetFormats(xml)
  } catch {
    return undefined
  }
}

/** the sidecar resolves every color to rgb; styles.xml still knows the theme slot and the pattern */
function withStoredColors(
  format: CellFormat | undefined,
  stored: StylesheetFormats | undefined,
  styleIndex: number,
  palette: readonly string[] | undefined,
): CellFormat | undefined {
  const xf = stored?.xfs[styleIndex]
  if (!stored || !xf) return format
  const out: CellFormat = { ...format }
  // the Normal font is usually theme="1": only a font the cell chose over it is a format
  const baseFont = stored.fontColors[stored.xfs[0]?.fontId ?? 0]
  const fontColor = stored.fontColors[xf.fontId]
  if (isThemeColor(fontColor) && !sameStyleColor(fontColor, baseFont)) {
    const echo = echoStyleColor(fontColor, palette)
    out.fontColor = echo.rgb
    out.fontColorTheme = { theme: echo.theme!, ...(echo.tint ? { tint: echo.tint } : {}) }
  }
  const fill = stored.fills[xf.fillId]
  if (fill) {
    // theme slots resolve here so fillColor and fill.fg.rgb agree byte for byte
    const display = fillDisplayColor(fill)
    if (display && (!out.fillColor || isThemeColor(display))) {
      out.fillColor = echoStyleColor(display, palette).rgb
    }
    if (isGradientFill(fill)) {
      const g = fill.gradient
      out.fill = {
        gradient: {
          ...(g.type ? { type: g.type } : {}),
          ...(g.angle !== undefined ? { angle: g.angle } : {}),
          ...(g.left !== undefined ? { left: g.left } : {}),
          ...(g.right !== undefined ? { right: g.right } : {}),
          ...(g.top !== undefined ? { top: g.top } : {}),
          ...(g.bottom !== undefined ? { bottom: g.bottom } : {}),
          stops: g.stops.map((stop) => ({
            position: stop.position,
            color: echoStyleColor(stop.color, palette),
          })),
        },
      }
    } else if (fill.pattern !== 'solid' || isThemeColor(fill.fg)) {
      out.fill = {
        pattern: fill.pattern,
        fg: echoStyleColor(fill.fg, palette),
        ...(fill.bg === undefined ? {} : { bg: echoStyleColor(fill.bg, palette) }),
      }
    }
  }
  return Object.keys(out).length ? out : undefined
}

function cellMatches(where: ReadWhere, value: Scalar, formula: string | undefined): boolean {
  switch (where) {
    case 'formula':
      return formula !== undefined
    case 'error':
      return typeof value === 'string' && ERROR_VALUE.test(value)
    case 'empty':
      return value === null && formula === undefined
    case 'number':
      return typeof value === 'number'
    case 'text':
      return typeof value === 'string' && !ERROR_VALUE.test(value)
  }
}

function sheetStats(
  wb: OpenedWorkbook,
  meta: SheetMeta,
  result: RangeResult,
  rows: Scalar[][],
  formulas: Record<string, string>,
  scanned: string,
): SheetStats {
  const stats: SheetStats = {
    sheets: wb.sheets.map((s) => ({ name: s.name, rows: s.rowCount, columns: s.columnCount })),
    usedRange:
      meta.rowCount && meta.columnCount
        ? `A1:${columnLabel(meta.columnCount - 1)}${meta.rowCount}`
        : null,
    rows: meta.rowCount,
    columns: meta.columnCount,
    scanned,
    nonEmpty: 0,
    formulas: Object.keys(formulas).length,
    errors: 0,
    numbers: 0,
    text: 0,
    booleans: 0,
    merges: result.merges?.length ?? 0,
  }
  for (const row of rows) {
    for (const v of row) {
      if (v === null) continue
      stats.nonEmpty++
      if (typeof v === 'number') stats.numbers++
      else if (typeof v === 'boolean') stats.booleans++
      else if (ERROR_VALUE.test(v)) stats.errors++
      else stats.text++
    }
  }
  return stats
}

/** OOXML column widths are in characters of the default font; 7 px per character is Calibri 11. */
const PX_PER_CHAR = 7

function areaLabel(a: CellArea): string {
  return `${columnLabel(a.startColumn)}${a.startRow + 1}:${columnLabel(a.endColumn)}${a.endRow + 1}`
}

/** Structure the values alone do not show: panes, filter, visuals and rule counts for the sheet; merges and links for the range. */
function sheetFeatures(wb: OpenedWorkbook, meta: SheetMeta, result: RangeResult): SheetFeatures {
  return {
    frozen: meta.freeze
      ? { rows: meta.freeze.frozenRows, columns: meta.freeze.frozenColumns }
      : null,
    filter: result.autoFilter ? areaLabel(result.autoFilter) : null,
    merges: (result.merges ?? []).map(areaLabel),
    charts: (wb.visuals ?? [])
      .filter((v) => v.sheetId === meta.id && v.kind === 'chart' && v.chart)
      .map((v) => ({
        id: v.id,
        title: v.chart!.title,
        types: v.chart!.chartTypes,
        series: v.chart!.series.length,
        anchor: `${columnLabel(v.anchor.fromColumn)}${v.anchor.fromRow + 1}`,
      })),
    tables: (meta.tables ?? []).map((t) => ({ name: t.name ?? null, range: areaLabel(t.range) })),
    pivots: (meta.pivotTables ?? []).map((p) => p.outputRef),
    conditionalFormats: result.conditionalRules?.length ?? 0,
    dataValidations: result.dataValidations?.length ?? 0,
    hyperlinks: result.hyperlinks?.length ?? 0,
    notes: meta.comments?.length ?? 0,
    hiddenColumns: (meta.columnWidths ?? [])
      .filter((w) => w.hidden)
      .reduce((n, w) => n + (w.endColumn - w.startColumn + 1), 0),
    hiddenSheet: meta.hidden === true,
  }
}

function cellFormat(style: CellStyle, base: CellStyle | undefined): CellFormat | undefined {
  const out: CellFormat = {}
  if (style.bold) out.bold = true
  if (style.italic) out.italic = true
  if (style.underline) out.underline = true
  if (style.strikethrough) out.strikethrough = true
  if (style.wrapText) out.wrapText = true
  if (style.fontFamily && style.fontFamily !== base?.fontFamily) out.fontFamily = style.fontFamily
  if (style.fontSize !== undefined && style.fontSize !== base?.fontSize)
    out.fontSize = style.fontSize
  if (style.fontColor) out.fontColor = style.fontColor
  if (style.fillColor) out.fillColor = style.fillColor
  if (style.horizontalAlignment) out.horizontalAlign = style.horizontalAlignment
  if (style.verticalAlignment) out.verticalAlign = style.verticalAlignment
  if (style.numberFormat && style.numberFormat !== 'General') out.numberFormat = style.numberFormat
  return Object.keys(out).length ? out : undefined
}

// ── csv ────────────────────────────────────────────────────────────────

/** The sidecar rejects a read_range above this many cells per request (lib.rs MAX_RANGE_CELLS). */
const RANGE_CELL_CAP = 100_000
/** Same ceiling as the app's Export as CSV (MAX_CSV_EXPORT_CHARS in apps/sheets/src/shared/ipc-channels.ts). */
const MAX_CSV_CHARS = 64_000_000

export interface SheetCsv {
  sheet: string
  sheets: number
  rows: number
  columns: number
  formulas: number
  csv: string
}

/**
 * One worksheet as CSV, cell text formatted the way the grid shows it (number
 * formats applied, formulas by their cached or freshly evaluated value); the
 * same shape the app's Export as CSV writes. Rows are produced band by band so
 * a wide, sparse used range never needs a dense grid of the whole sheet.
 */
export async function sheetToCsv(path: string, opts: { sheet?: string }): Promise<SheetCsv> {
  return withOpenWorkbook(path, async (client, wb) => {
    const meta = pickSheet(wb, opts.sheet) as SheetMeta
    const styles = wb.styles ?? []
    const date1904 = wb.date1904 === true
    const columnStyles = new Map<number, number>()
    for (const span of meta.columnWidths ?? []) {
      if (span.styleIndex === undefined) continue
      for (let c = span.startColumn; c <= span.endColumn && c < meta.columnCount; c++) {
        columnStyles.set(c, span.styleIndex)
      }
    }
    const lines: string[] = []
    let chars = 0
    let formulas = 0
    if (meta.rowCount > 0 && meta.columnCount > 0) {
      const whole = {
        startRow: 0,
        endRow: meta.rowCount - 1,
        startColumn: 0,
        endColumn: meta.columnCount - 1,
      }
      for (const band of recalcBands(whole, RANGE_CELL_CAP)) {
        const r = await readRangeIndexed(client, wb.sessionId, meta.id, band)
        const height = band.endRow - band.startRow + 1
        const grid: string[][] = Array.from({ length: height }, () =>
          Array<string>(meta.columnCount).fill(''),
        )
        // OOXML precedence: the cell's own xf, then the row default, then the column default
        const rowStyles = new Map((r.rows ?? []).map((row) => [row.row, row.styleIndex]))
        const pending: { row: number; column: number }[] = []
        for (const c of r.cells) {
          if (c.formula) formulas++
          if (c.value === undefined || c.value === null) {
            if (c.formula) pending.push({ row: c.row, column: c.column })
            continue
          }
          const styleIndex = c.styleIndex ?? rowStyles.get(c.row) ?? columnStyles.get(c.column)
          const format = styleIndex === undefined ? undefined : styles[styleIndex]?.numberFormat
          grid[c.row - band.startRow]![c.column] = displayText(c.value, format, date1904)
        }
        if (pending.length) {
          for (const cell of await recalcRange(client, path, meta.name, boundingBox(pending))) {
            const row = grid[cell.row - band.startRow]
            if (row && row[cell.column] === '') row[cell.column] = cell.formatted
          }
        }
        for (const row of grid) {
          const line = row.map(csvField).join(',') + '\r\n'
          chars += line.length
          if (chars > MAX_CSV_CHARS) {
            throw new CliError(EXIT.conversion, 'sheet is too large to export as CSV', {
              limit_chars: MAX_CSV_CHARS,
            })
          }
          lines.push(line)
        }
      }
    }
    return {
      sheet: meta.name,
      sheets: wb.sheets.length,
      rows: meta.rowCount,
      columns: meta.columnCount,
      formulas,
      csv: lines.join(''),
    }
  })
}

/** Grid text for one value; `format` is the cell's resolved number format. */
export function displayText(value: Scalar, format: string | undefined, date1904: boolean): string {
  if (typeof value === 'string') return value
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  if (value === null) return ''
  const pattern = format || 'General'
  // 1904-system workbooks store calendar dates 1462 days lower; time-only and
  // elapsed patterns ([h]:mm) show the serial's magnitude and must not shift
  const serial = date1904 && isCalendarDatePattern(pattern) ? value + 1462 : value
  return numfmt.format(pattern, serial, { throws: false })
}

function isCalendarDatePattern(pattern: string): boolean {
  try {
    const type = (numfmt.getFormatInfo(pattern) as { type?: string }).type
    return type === 'date' || type === 'datetime'
  } catch {
    return false
  }
}

// same quoting as the app's csv-export.ts (RFC 4180: quote on comma, quote or newline)
function csvField(text: string): string {
  const normalized = text.replace(/\r\n|\r/g, '\n')
  return /[",\n]/.test(normalized) ? `"${normalized.replace(/"/g, '""')}"` : normalized
}

// ── writing ────────────────────────────────────────────────────────────

export interface CellInput {
  sheet?: string
  cell: string
  value?: Scalar
  formula?: string
  style?: WorkbookStyleEdit
}

export interface TableInput {
  name: string
  rows: Scalar[][]
  /** `row:column` → number format for cells whose value was typed on import (ISO dates) */
  formats?: Record<string, string>
}

export interface CsvTableOptions {
  /** decimal separator used by the file; `,` also accepts `.` as a thousands separator */
  decimal?: '.' | ','
}

const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30)
const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?$/

/** Serial day number Excel stores for an ISO date or date-time; undefined for anything else. */
export function isoDateSerial(value: string): { serial: number; format: string } | undefined {
  const m = ISO_DATE.exec(value.trim())
  if (!m) return undefined
  const [y, mo, d, hh, mm, ss] = m
    .slice(1)
    .map((part) => (part === undefined ? undefined : Number(part)))
  const utc = Date.UTC(y!, mo! - 1, d!, hh ?? 0, mm ?? 0, ss ?? 0)
  const date = new Date(utc)
  if (date.getUTCFullYear() !== y || date.getUTCMonth() !== mo! - 1 || date.getUTCDate() !== d) {
    return undefined
  }
  if ((hh ?? 0) > 23 || (mm ?? 0) > 59 || (ss ?? 0) > 59) return undefined
  const serial = (utc - EXCEL_EPOCH_UTC) / 86_400_000
  const format =
    hh === undefined ? 'yyyy-mm-dd' : ss === undefined ? 'yyyy-mm-dd hh:mm' : 'yyyy-mm-dd hh:mm:ss'
  return { serial, format }
}

/** `1.234,5` → 1234.5 under a comma decimal; plain `1234,5` too. Leading zeros stay text as in isNumericCell. */
function commaDecimalNumber(value: string): number | undefined {
  if (!/^-?(?:\d{1,3}(?:\.\d{3})+|\d+)(?:,\d+)?$/.test(value)) return undefined
  const normalized = value.replace(/\./g, '').replace(',', '.')
  return isNumericCell(normalized) ? Number(normalized) : undefined
}

const STYLES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
  '</styleSheet>'

const STYLES_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
const STYLES_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'

/**
 * The blank workbook normally ships a stylesheet of its own; this tops one up for a
 * blank that does not, because the formula engine refuses to import such a package.
 * Every part is added only when missing — a second styles Override or Relationship
 * would duplicate a PartName and a Relationship Id, which makes Excel repair the file.
 */
export async function blankWorkbook(sheetName = 'Sheet1'): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await blankXlsxBuffer(sheetName))
  if (!zip.file('xl/styles.xml')) zip.file('xl/styles.xml', STYLES_XML)
  const types = await zip.file('[Content_Types].xml')!.async('string')
  if (!types.includes('PartName="/xl/styles.xml"')) {
    zip.file(
      '[Content_Types].xml',
      types.replace(
        '</Types>',
        `<Override PartName="/xl/styles.xml" ContentType="${STYLES_CONTENT_TYPE}"/></Types>`,
      ),
    )
  }
  const rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string')
  if (!rels.includes(`Type="${STYLES_REL_TYPE}"`)) {
    zip.file(
      'xl/_rels/workbook.xml.rels',
      rels.replace(
        '</Relationships>',
        `<Relationship Id="${nextFreeRelationshipId(rels)}" ` +
          `Type="${STYLES_REL_TYPE}" Target="styles.xml"/></Relationships>`,
      ),
    )
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

/** CSV → table; numeric-looking cells become numbers like the app's importer. */
export function csvTable(bytes: Uint8Array, name: string, opts: CsvTableOptions = {}): TableInput {
  const formats: Record<string, string> = {}
  const rows = parseCsv(decodeCsvBuffer(bytes)).map((r, rowIndex) =>
    r.map((v, columnIndex) => {
      if (opts.decimal === ',') {
        const n = commaDecimalNumber(v)
        if (n !== undefined) return n
      } else if (isNumericCell(v)) {
        return Number(v)
      }
      const date = isoDateSerial(v)
      if (!date) return v
      formats[`${rowIndex}:${columnIndex}`] = date.format
      return date.serial
    }),
  )
  return Object.keys(formats).length ? { name, rows, formats } : { name, rows }
}

export function cellEditsFromTable(
  sheet: string,
  rows: Scalar[][],
  formats: Record<string, string> = {},
): CellEdit[] {
  const edits: CellEdit[] = []
  rows.forEach((row, r) => {
    row.forEach((raw, c) => {
      if (raw === null || raw === undefined || raw === '') return
      const numberFormat = formats[`${r}:${c}`]
      edits.push(cellEdit(sheet, r, c, raw, numberFormat ? { numberFormat } : undefined))
    })
  })
  return edits
}

function cellEdit(
  sheet: string,
  row: number,
  column: number,
  raw: Scalar,
  style?: WorkbookStyleEdit,
): CellEdit {
  const formula = typeof raw === 'string' && raw.startsWith('=') ? raw : undefined
  return {
    sheetName: sheet,
    row,
    column,
    writeValue: true,
    cell: formula ? { value: null, formula } : { value: raw },
    ...(style ? { style } : {}),
  }
}

export function cellEditsFromInputs(inputs: CellInput[], defaultSheet: string): CellEdit[] {
  return inputs.map((input, i) => {
    if (typeof input.cell !== 'string') {
      throw new CliError(EXIT.usage, `cells[${i}]: missing "cell" address (e.g. "B2")`, undefined, {
        reason: 'invalid_argument',
      })
    }
    let coords: { row: number; column: number }
    try {
      coords = parseAddressChecked(input.cell)
    } catch {
      throw new CliError(EXIT.usage, `cells[${i}]: invalid address "${input.cell}"`, undefined, {
        reason: 'invalid_argument',
      })
    }
    const raw: Scalar =
      input.formula !== undefined
        ? input.formula.startsWith('=')
          ? input.formula
          : `=${input.formula}`
        : (input.value ?? null)
    const styleOnly = input.value === undefined && input.formula === undefined
    const edit = cellEdit(input.sheet ?? defaultSheet, coords.row, coords.column, raw, input.style)
    return styleOnly ? { ...edit, writeValue: false } : edit
  })
}

export interface WriteOutcome {
  cells: number
  formulas: number
  /** false when formulas were written but the engine could not fill their cached values */
  cachedValues: boolean
  /** `Sheet!A1` of formulas left without a cached value on purpose */
  uncached?: string[]
  /** `Sheet!A1` of formulas the engine could not parse (a reference it cannot read) */
  unparsed?: string[]
  warning?: string
  /** formulas the engine could not parse — reported apart from `warning`, which is about missing functions */
  formulaError?: string
  /** side effects a program should know about that are not formula related */
  notes?: { code: string; message: string }[]
}

/**
 * Writes the edits, then evaluates the formulas on the written file and
 * re-applies the same edits with the results as cached values, so readers
 * that trust <v> (openpyxl, pandas, quick previews) see numbers too. The
 * file is complete after the first write; the cached-value pass is best
 * effort and never turns a finished write into a failure.
 */
export async function writeWorkbook(
  source: Buffer,
  edits: readonly CellEdit[],
  outputPath: string,
  opts: {
    plan?: SheetEditPlan
    structuralOps?: readonly SheetStructuralOps[]
    /** original → new sheet names applied by `plan`; the written file carries the new ones */
    renames?: Record<string, string>
    /** the declarative save payloads of the DSL ops the in-memory workbook cannot hold */
    gateway?: GatewayPayloads
  } = {},
): Promise<WriteOutcome> {
  const { plan, structuralOps = [], renames = {}, gateway = EMPTY_PAYLOADS } = opts
  const notes: { code: string; message: string }[] = []
  if (edits.some((e) => e.style && styleHasThemeColor(e.style)) && !(await hasThemePart(source))) {
    // without xl/theme/theme1.xml Excel has no palette to resolve a slot against
    edits = edits.map((e) => (e.style ? { ...e, style: flattenThemeColors(e.style) } : e))
    notes.push({
      code: 'theme_colors_flattened',
      message:
        'the workbook has no theme part, so theme colors were written as their Office-theme rgb values',
    })
  }
  const save = async (formulaValues: readonly SheetFormulaValues[]): Promise<XlsxMutation> => {
    const mutation = await planCellEditsToXlsx(
      await createBufferEntrySource(source),
      edits,
      structuralOps,
      gateway.chartEdits,
      plan,
      gateway.filterStates,
      gateway.hyperlinkEdits,
      gateway.cfStates,
      gateway.dvStates,
      gateway.sheetProtections,
      gateway.definedNamesState,
      gateway.visualAdditions,
      gateway.pageSetupStates,
      gateway.noteStates,
      gateway.tableAdditions,
      gateway.pivotAdditions,
      [],
      [],
      [],
      gateway.sparklineAdditions,
      formulaValues,
    )
    return assembleWithJsZip(source, mutation)
  }
  let first: XlsxMutation
  try {
    first = await save([])
  } catch (err) {
    // the gateway fails closed with a sentence about the file (x14 rules, name clashes, table overlaps)
    throw new CliError(
      EXIT.usage,
      `ops rejected by the workbook writer: ${(err as Error).message}`,
      undefined,
      { reason: 'op_rejected' },
    )
  }
  await atomicWriteFile(outputPath, first.buffer)
  const formulaCells = edits.filter((e) => e.cell.formula)
  const base = {
    cells: edits.length,
    formulas: formulaCells.length,
    ...(notes.length ? { notes } : {}),
  }
  if (formulaCells.length === 0) return { ...base, cachedValues: true }
  if (!xlsxSidecarPath()) {
    return {
      ...base,
      cachedValues: false,
      warning: 'xlsx engine not found; formulas recalculate on open',
    }
  }
  try {
    const { values, uncached, unparsed } = await evaluateFormulas(outputPath, formulaCells, renames)
    if (values.length) {
      const second = await save(values)
      await atomicWriteFile(outputPath, second.buffer)
      if (uncached.length === 0 && unparsed.length === 0) return { ...base, cachedValues: true }
      const skipped = uncached.length + unparsed.length
      return {
        ...base,
        cachedValues: skipped < formulaCells.length,
        ...(uncached.length ? { uncached } : {}),
        ...(unparsed.length ? { unparsed } : {}),
        ...(uncached.length
          ? {
              warning: `${uncached.length} formula(s) use functions the local engine does not evaluate (${list(uncached)}); they have no cached value and recalculate on open`,
            }
          : {}),
        ...(unparsed.length
          ? {
              formulaError: `${unparsed.length} formula(s) the local engine could not parse (${list(unparsed)}); check their references — they have no cached value and recalculate on open`,
            }
          : {}),
      }
    }
    return {
      ...base,
      cachedValues: false,
      warning:
        'the engine returned no results for the written formulas; the file recalculates on open',
    }
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err)
    return {
      ...base,
      cachedValues: false,
      warning: `formula results not cached (${message}); the file is complete and recalculates on open`,
    }
  }
}

/**
 * `#NAME?` from the local engine means "a function it does not implement" at
 * least as often as a typo, and `#ERROR!` is the engine's own failure, not an
 * Excel error: neither result is cached. The cell keeps its formula and no
 * <v>, and Excel computes it on open (fullCalcOnLoad).
 */
/** The engine knows the syntax but not the function: Excel computes it on open. */
const UNKNOWN_FUNCTION_RESULT = '#NAME?'
/** The engine could not parse the formula at all — a reference it cannot read, not a missing function. */
const UNPARSED_RESULT = '#ERROR!'
const list = (cells: readonly string[]) =>
  cells.slice(0, 3).join(', ') + (cells.length > 3 ? ', …' : '')
async function evaluateFormulas(
  path: string,
  formulaCells: readonly CellEdit[],
  renames: Record<string, string>,
): Promise<{ values: SheetFormulaValues[]; uncached: string[]; unparsed: string[] }> {
  const bySheet = new Map<string, CellEdit[]>()
  for (const e of formulaCells) bySheet.set(e.sheetName, [...(bySheet.get(e.sheetName) ?? []), e])
  // edits carry the file's original sheet names; the written file has the renamed ones
  const wanted = new Set(formulaCells.map((c) => `${c.sheetName} ${c.row} ${c.column}`))
  const originalName = (written: string) =>
    Object.entries(renames).find(([, after]) => after === written)?.[0] ?? written
  return withSidecar(async (client) => {
    const out: SheetFormulaValues[] = []
    const uncached: string[] = []
    const unparsed: string[] = []
    for (const [sheet, cells] of bySheet) {
      const evaluated = await recalcRange(client, path, renames[sheet] ?? sheet, boundingBox(cells))
      const values = evaluated
        .filter((cell) => wanted.has(`${originalName(cell.sheet)} ${cell.row} ${cell.column}`))
        .map((cell) => {
          // Both leave the cell without a cached value — the file is correct
          // and Excel recomputes on open — but they have different causes and
          // the caller reports them apart.
          if (cell.formatted === UNKNOWN_FUNCTION_RESULT) {
            uncached.push(`${sheet}!${toA1Address(cell.row, cell.column)}`)
            return { row: cell.row, column: cell.column, value: null }
          }
          if (cell.formatted === UNPARSED_RESULT) {
            unparsed.push(`${sheet}!${toA1Address(cell.row, cell.column)}`)
            return { row: cell.row, column: cell.column, value: null }
          }
          if (cell.isError) {
            return { row: cell.row, column: cell.column, value: { error: cell.formatted } }
          }
          return {
            row: cell.row,
            column: cell.column,
            value: (cell.number ?? (cell.formatted === '' ? null : cell.formatted)) as Scalar,
          }
        })
      // the refresh is keyed like the edits, by the file's original sheet name
      if (values.length) out.push({ sheetName: sheet, cells: values })
    }
    return { values: out, uncached, unparsed }
  })
}

interface Bounds {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}

function boundingBox(cells: readonly { row: number; column: number }[]): Bounds {
  // a loop, not Math.min(...spread): a large uncached window would overflow the call stack
  const box = { startRow: Infinity, endRow: -Infinity, startColumn: Infinity, endColumn: -Infinity }
  for (const c of cells) {
    if (c.row < box.startRow) box.startRow = c.row
    if (c.row > box.endRow) box.endRow = c.row
    if (c.column < box.startColumn) box.startColumn = c.column
    if (c.column > box.endColumn) box.endColumn = c.column
  }
  return box
}

/** The sidecar rejects a recalc read above this many cells per request (recalc.rs MAX_RECALC_READ_CELLS). */
export const RECALC_CELL_CAP = 20_000

/** Splits a rectangle into row (and, for very wide ranges, column) bands that each fit the cap. */
export function recalcBands(bounds: Bounds, cap = RECALC_CELL_CAP): Bounds[] {
  const width = bounds.endColumn - bounds.startColumn + 1
  const bands: Bounds[] = []
  const colStep = Math.min(width, cap)
  for (let c0 = bounds.startColumn; c0 <= bounds.endColumn; c0 += colStep) {
    const c1 = Math.min(bounds.endColumn, c0 + colStep - 1)
    const rowStep = Math.max(1, Math.floor(cap / (c1 - c0 + 1)))
    for (let r0 = bounds.startRow; r0 <= bounds.endRow; r0 += rowStep) {
      bands.push({
        startRow: r0,
        endRow: Math.min(bounds.endRow, r0 + rowStep - 1),
        startColumn: c0,
        endColumn: c1,
      })
    }
  }
  return bands
}

async function recalcRange(
  client: XlsxSidecarClient,
  path: string,
  sheet: string,
  bounds: Bounds,
): Promise<RecalcResult['cells']> {
  const cells: RecalcResult['cells'] = []
  // one band per request: the recalc worker is single-flight
  for (const range of recalcBands(bounds)) {
    const r = (await client.recalcCells({
      path,
      edits: [],
      reads: [{ sheet, range }],
    })) as RecalcResult
    cells.push(...r.cells)
  }
  return cells
}

/** What the workbook engine computes for a range of the file on disk, keyed `row|column` (0-based). */
export async function computedValues(
  path: string,
  sheet: string,
  bounds: Bounds,
): Promise<Map<string, { value: Scalar; isError: boolean }>> {
  const out = new Map<string, { value: Scalar; isError: boolean }>()
  await withSidecar(async (client) => {
    for (const cell of await recalcRange(client, path, sheet, bounds)) {
      out.set(`${cell.row}|${cell.column}`, {
        value: cell.number ?? (cell.formatted === '' ? null : cell.formatted),
        isError: cell.isError === true,
      })
    }
  })
  return out
}

/** Multi-sheet plan for a fresh workbook whose blank already carries `first`. */
export function additionPlan(first: string, others: readonly string[]): SheetEditPlan | undefined {
  if (others.length === 0) return undefined
  return {
    renames: [],
    additions: others.map((name) => ({ name })),
    removals: [],
    order: [first, ...others],
  }
}
