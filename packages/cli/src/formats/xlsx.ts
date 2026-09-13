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
  type CellEdit,
  type SheetFormulaValues,
  type SheetStructuralOps,
  type XlsxMutation,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { SheetEditPlan } from '@genoffice/xlsx-gateway/gateway/xlsx-sheets'
import { EMPTY_PAYLOADS, type GatewayPayloads } from './xlsx-gateway-ops'
import type { WorkbookStyleEdit } from '@genoffice/xlsx-gateway/shared/edit-schemas'
import { atomicWriteFile } from '../../../../apps/sheets/src/main/atomic-write'
import { XlsxSidecarClient } from '../../../../apps/sheets/src/main/xlsx-sidecar-client'
import { CliError, EXIT } from '../result'
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
  }[]
  activeTab: number
  definedNames?: unknown[]
  styles?: CellStyle[]
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
    throw new CliError(EXIT.usage, `sheet not found: ${name}`, {
      sheets: wb.sheets.map((s) => s.name),
    })
  }
  return meta
}

interface RecalcResult {
  cells: { sheet: string; row: number; column: number; formatted: string; number?: number }[]
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
  fillColor?: string
  horizontalAlign?: string
  verticalAlign?: string
  numberFormat?: string
}

export interface SheetRead {
  sheet: string
  range: string
  rows: Scalar[][]
  formulas: Record<string, string>
  truncated: boolean
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
    throw new CliError(EXIT.usage, `invalid ${what} address: ${address}`)
  }
  if (parsed.row >= MAX_ROWS || parsed.column >= MAX_COLUMNS) {
    throw new CliError(
      EXIT.usage,
      `${what} address outside the sheet limits (XFD1048576): ${address}`,
    )
  }
  return parsed
}

function parseRangeChecked(range: string): Bounds {
  const parts = range.split(':')
  if (parts.length > 2 || !parts[0]) throw new CliError(EXIT.usage, `invalid range: ${range}`)
  const first = parseAddressChecked(parts[0], 'range')
  const second = parts[1] ? parseAddressChecked(parts[1], 'range') : first
  return {
    startRow: Math.min(first.row, second.row),
    startColumn: Math.min(first.column, second.column),
    endRow: Math.max(first.row, second.row),
    endColumn: Math.max(first.column, second.column),
  }
}

export async function readSheet(
  path: string,
  opts: { sheet?: string; range?: string; formats?: boolean },
): Promise<SheetRead> {
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
    const read: SheetRead = {
      sheet: meta.name,
      range: rangeLabel,
      rows,
      formulas,
      truncated,
      features: sheetFeatures(wb, meta, result),
    }
    if (opts.formats) {
      const formats: Record<string, CellFormat> = {}
      // xf 0 is the workbook default; a style only counts where it differs from it
      const base = wb.styles?.[0]
      for (const c of result.cells) {
        if (!c.styleIndex) continue
        const style = wb.styles?.[c.styleIndex]
        const format = style ? cellFormat(style, base) : undefined
        if (format) formats[`${columnLabel(c.column)}${c.row + 1}`] = format
      }
      read.formats = formats
      read.columnWidths = Object.fromEntries(
        (meta.columnWidths ?? [])
          .filter((w) => w.width !== undefined)
          .flatMap((w) => {
            const out: [string, number][] = []
            for (let c = w.startColumn; c <= Math.min(w.endColumn, bounds.endColumn); c++) {
              if (c >= bounds.startColumn)
                out.push([columnLabel(c), Math.round(w.width! * PX_PER_CHAR)])
            }
            return out
          }),
      )
      read.rowHeights = Object.fromEntries(
        (result.rows ?? [])
          .filter((r) => r.height !== undefined)
          .map((r) => [String(r.row + 1), r.height!]),
      )
    }
    return read
  })
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

/** The app's blank workbook has no stylesheet; the formula engine refuses to import such a package, so add a minimal one. */
export async function blankWorkbook(sheetName = 'Sheet1'): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await blankXlsxBuffer(sheetName))
  zip.file('xl/styles.xml', STYLES_XML)
  const types = await zip.file('[Content_Types].xml')!.async('string')
  zip.file(
    '[Content_Types].xml',
    types.replace(
      '</Types>',
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
    ),
  )
  const rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string')
  zip.file(
    'xl/_rels/workbook.xml.rels',
    rels.replace(
      '</Relationships>',
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

/** CSV → table; numeric-looking cells become numbers like the app's importer. */
export function csvTable(bytes: Uint8Array, name: string): TableInput {
  const rows = parseCsv(decodeCsvBuffer(bytes)).map((r) =>
    r.map((v) => (isNumericCell(v) ? Number(v) : v)),
  )
  return { name, rows }
}

export function cellEditsFromTable(sheet: string, rows: Scalar[][]): CellEdit[] {
  const edits: CellEdit[] = []
  rows.forEach((row, r) => {
    row.forEach((raw, c) => {
      if (raw === null || raw === undefined || raw === '') return
      edits.push(cellEdit(sheet, r, c, raw))
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
      throw new CliError(EXIT.usage, `cells[${i}]: missing "cell" address (e.g. "B2")`)
    }
    let coords: { row: number; column: number }
    try {
      coords = parseAddressChecked(input.cell)
    } catch {
      throw new CliError(EXIT.usage, `cells[${i}]: invalid address "${input.cell}"`)
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
  warning?: string
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
      [],
      [],
      [],
      [],
      [],
      formulaValues,
    )
    return assembleWithJsZip(source, mutation)
  }
  let first: XlsxMutation
  try {
    first = await save([])
  } catch (err) {
    // the gateway fails closed with a sentence about the file (x14 rules, name clashes, table overlaps)
    throw new CliError(EXIT.usage, `ops rejected by the workbook writer: ${(err as Error).message}`)
  }
  await atomicWriteFile(outputPath, first.buffer)
  const formulaCells = edits.filter((e) => e.cell.formula)
  const base = { cells: edits.length, formulas: formulaCells.length }
  if (formulaCells.length === 0) return { ...base, cachedValues: true }
  if (!xlsxSidecarPath()) {
    return {
      ...base,
      cachedValues: false,
      warning: 'xlsx engine not found; formulas recalculate on open',
    }
  }
  try {
    const values = await evaluateFormulas(outputPath, formulaCells, renames)
    if (values.length) {
      const second = await save(values)
      await atomicWriteFile(outputPath, second.buffer)
      return { ...base, cachedValues: true }
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

async function evaluateFormulas(
  path: string,
  formulaCells: readonly CellEdit[],
  renames: Record<string, string>,
): Promise<SheetFormulaValues[]> {
  const bySheet = new Map<string, CellEdit[]>()
  for (const e of formulaCells) bySheet.set(e.sheetName, [...(bySheet.get(e.sheetName) ?? []), e])
  // edits carry the file's original sheet names; the written file has the renamed ones
  const wanted = new Set(formulaCells.map((c) => `${c.sheetName} ${c.row} ${c.column}`))
  const originalName = (written: string) =>
    Object.entries(renames).find(([, after]) => after === written)?.[0] ?? written
  return withSidecar(async (client) => {
    const out: SheetFormulaValues[] = []
    for (const [sheet, cells] of bySheet) {
      const evaluated = await recalcRange(client, path, renames[sheet] ?? sheet, boundingBox(cells))
      const values = evaluated
        .filter((cell) => wanted.has(`${originalName(cell.sheet)} ${cell.row} ${cell.column}`))
        .map((cell) => ({
          row: cell.row,
          column: cell.column,
          value: (cell.number ?? (cell.formatted === '' ? null : cell.formatted)) as Scalar,
        }))
      // the refresh is keyed like the edits, by the file's original sheet name
      if (values.length) out.push({ sheetName: sheet, cells: values })
    }
    return out
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
