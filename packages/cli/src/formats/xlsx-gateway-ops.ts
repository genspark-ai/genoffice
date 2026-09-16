import { existsSync, readFileSync, statSync } from 'node:fs'
import { homedir } from 'node:os'
import { isAbsolute, resolve } from 'node:path'
import { fetchRemoteImage } from '@genoffice/electron-utils/remote-image'
import {
  columnIndex,
  columnLabel,
  parseAddress,
  parseRange,
  type RangeBounds,
} from '@genoffice/xlsx-gateway/domain/cell-address'
import type { ChartVisualState } from '@genoffice/xlsx-gateway/domain/chart-visual'
import {
  matchableCellText,
  type WorkbookOperation,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import type {
  WorkbookSnapshot,
  WorksheetState,
} from '@genoffice/xlsx-gateway/domain/workbook.types'
import type { CfWireRule } from '@genoffice/xlsx-gateway/gateway/xlsx-cf'
import type {
  DefinedNameEntry,
  DefinedNamesState,
} from '@genoffice/xlsx-gateway/gateway/xlsx-defined-names'
import type { ChartAdd, DrawingAnchor } from '@genoffice/xlsx-gateway/gateway/xlsx-drawing-add'
import type { DvWireRule } from '@genoffice/xlsx-gateway/gateway/xlsx-dv'
import {
  areasOverlap,
  buildPivotLayout,
  PivotLayoutError,
  pivotOutputArea,
  type PivotScalar,
} from '@genoffice/xlsx-gateway/domain/pivot-layout'
import type { SheetFilterState } from '@genoffice/xlsx-gateway/gateway/xlsx-filter'
import type {
  SheetCfState,
  SheetDvState,
  SheetHyperlinkEdits,
  SheetNoteState,
  SheetPivotAddition,
  SheetProtectionState,
  SheetSparklineAddition,
  SheetStructuralOps,
  SheetTableAddition,
  SheetVisualAddition,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { SheetNote } from '@genoffice/xlsx-gateway/gateway/xlsx-notes'
import type { SheetPageSetupState } from '@genoffice/xlsx-gateway/gateway/xlsx-page-setup'
import type { WorkbookChartEdit } from '@genoffice/xlsx-gateway/shared/edit-schemas'
import { assertAllowed, type PathContext } from '../fs'
import { classifyOpError } from '../op-errors'
import { CliError, EXIT } from '../result'
import { imageSize } from './image-size'

/**
 * Ops the in-memory workbook validates but cannot apply (its snapshot has
 * nowhere to keep links, filters, rules, visuals or sheet visibility). The
 * gateway writes all of them into the file; this module turns the ops into the
 * gateway's declarative save payloads, the way the app's edit journal does.
 */
export const GATEWAY_DSL_OPS = [
  'add_chart',
  'edit_chart',
  'add_image',
  'add_shape',
  'add_table',
  'set_hyperlink',
  'set_note',
  'set_freeze',
  'set_page_setup',
  'set_rows_hidden',
  'set_cols_hidden',
  'set_sheet_hidden',
  'move_sheet',
  'duplicate_sheet',
  'protect_sheet',
  'set_filter',
  'clear_filter',
  'set_filter_criteria',
  'add_conditional_format',
  'clear_conditional_formats',
  'set_data_validation',
  'add_defined_name',
  'delete_defined_name',
  'add_pivot',
  'add_sparkline',
] as const

/** Ops whose save payload only the live editor can produce, with the reason the agent sees. */
export const REFUSED_DSL_OPS: Readonly<Record<string, string>> = {
  refresh_pivot: 'the pivot layout is computed by the editor',
  add_table_row: 'edits a table created in the same editor session',
  add_table_column: 'edits a table created in the same editor session',
  delete_table_row: 'edits a table created in the same editor session',
  delete_table_column: 'edits a table created in the same editor session',
  delete_table: 'removes a table created in the same editor session',
  edit_shape: 'targets a shape created in the same editor session',
  delete_visual: 'targets a visual by its editor session id',
}

/** Existing per-sheet state the payloads must carry over (the gateway rewrites whole sections). */
export interface SheetFileState {
  conditionalFormats: number
  dataValidations: number
  autoFilter: { range: RangeBounds; hasCriteria: boolean } | null
  /** 0-based rows the file hides */
  hiddenRows: number[]
  notes: SheetNote[]
  pivots: { name: string; range: RangeBounds }[]
}

/** A pivot output cell the batch writes next to the pivot definition. */
export interface BakedCell {
  sheetName: string
  row: number
  column: number
  value: string | number
  numberFormat?: string
}

export interface WorkbookFileState {
  /** file sheet names in tab order */
  order: string[]
  definedNames: DefinedNameEntry[]
  /** table names are workbook-scoped */
  tableNames: string[]
  sheets: Map<string, SheetFileState>
}

export interface SheetTabOps {
  moves: { name: string; position: number }[]
  hidden: { name: string; hidden: boolean }[]
  duplicates: { source: string; name: string }[]
}

export interface GatewayPayloads {
  hyperlinkEdits: SheetHyperlinkEdits[]
  noteStates: SheetNoteState[]
  pageSetupStates: SheetPageSetupState[]
  filterStates: SheetFilterState[]
  cfStates: SheetCfState[]
  dvStates: SheetDvState[]
  definedNamesState: DefinedNamesState | null
  visualAdditions: SheetVisualAddition[]
  sheetProtections: SheetProtectionState[]
  tableAdditions: SheetTableAddition[]
  chartEdits: WorkbookChartEdit[]
  hiddenOps: SheetStructuralOps[]
  tabs: SheetTabOps
  pivotAdditions: SheetPivotAddition[]
  bakedCells: BakedCell[]
  sparklineAdditions: SheetSparklineAddition[]
  warnings: string[]
}

export const EMPTY_PAYLOADS: GatewayPayloads = {
  hyperlinkEdits: [],
  noteStates: [],
  pageSetupStates: [],
  filterStates: [],
  cfStates: [],
  dvStates: [],
  definedNamesState: null,
  visualAdditions: [],
  sheetProtections: [],
  tableAdditions: [],
  chartEdits: [],
  hiddenOps: [],
  tabs: { moves: [], hidden: [], duplicates: [] },
  pivotAdditions: [],
  bakedCells: [],
  sparklineAdditions: [],
  warnings: [],
}

export function hasGatewayPayloads(p: GatewayPayloads): boolean {
  return (
    p.hyperlinkEdits.length +
      p.noteStates.length +
      p.pageSetupStates.length +
      p.filterStates.length +
      p.cfStates.length +
      p.dvStates.length +
      p.visualAdditions.length +
      p.sheetProtections.length +
      p.tableAdditions.length +
      p.chartEdits.length +
      p.hiddenOps.length +
      p.tabs.moves.length +
      p.tabs.hidden.length +
      p.tabs.duplicates.length +
      p.pivotAdditions.length +
      p.sparklineAdditions.length >
      0 || p.definedNamesState !== null
  )
}

export interface GatewayBuildInput {
  /** the batch, normalized (sheet names resolved to ids), in order; index = position in the batch */
  ops: readonly (WorkbookOperation & { readonly __index: number })[]
  /** sheet id → the name the file uses */
  namesById: Map<string, string>
  /** the workbook as read from the file, before the batch */
  before: WorkbookSnapshot
  /** the workbook after the adapter applied what it could */
  after: WorkbookSnapshot
  file: WorkbookFileState
  ctx: PathContext | undefined
  /** formula results the workbook engine computes from the file on disk; absent when the batch runs on a buffer */
  computed?: (
    sheetName: string,
    bounds: RangeBounds,
  ) => Promise<Map<string, { value: PivotScalar; isError: boolean }>>
}

const MAX_SPARKLINES_PER_OP = 200

const NOTE_AUTHOR = 'GenOffice'
const MAX_IMAGE_BYTES = 20 * 1024 * 1024

type FilterDraft = { range: RangeBounds; columns: Map<number, string[]>; cleared: boolean }

export async function buildGatewayPayloads(input: GatewayBuildInput): Promise<GatewayPayloads> {
  const { ops, namesById, before, after, file } = input
  const name = (sheetId: string) => namesById.get(sheetId) ?? sheetId
  const fileState = (sheetName: string): SheetFileState =>
    file.sheets.get(sheetName) ?? {
      conditionalFormats: 0,
      dataValidations: 0,
      autoFilter: null,
      hiddenRows: [],
      notes: [],
      pivots: [],
    }
  const sheetAfter = (sheetId: string): WorksheetState | undefined =>
    after.sheets.find((s) => s.id === sheetId)
  const reject = (index: number, op: string, message: string): never => {
    throw new CliError(
      EXIT.usage,
      `ops[${index}] (${op}) rejected: ${message}`,
      { failures: [classifyOpError(index, op, message)] },
      {
        reason: 'op_rejected',
        suggestion: 'fix the op against `genoffice guide sheets`, then resend the whole batch',
      },
    )
  }

  const hyperlinks = new Map<string, SheetHyperlinkEdits['edits'][number][]>()
  const notes = new Map<string, SheetNote[]>()
  type PageSetupPatch = {
    -readonly [K in keyof Omit<SheetPageSetupState, 'sheetName'>]: SheetPageSetupState[K]
  }
  const pageSetup = new Map<string, PageSetupPatch>()
  const filters = new Map<string, FilterDraft>()
  const cf = new Map<string, { rules: CfWireRule[]; cleared: boolean }>()
  const dv = new Map<string, { rules: DvWireRule[]; remove: RangeBounds[] }>()
  const protections = new Map<string, boolean>()
  const hidden = new Map<string, SheetStructuralOps['ops'][number][]>()
  const tables: SheetTableAddition[] = []
  const visuals: SheetVisualAddition[] = []
  const chartEdits: WorkbookChartEdit[] = []
  const pivots: SheetPivotAddition[] = []
  const baked: BakedCell[] = []
  const sparklines: SheetSparklineAddition[] = []
  const tabs: SheetTabOps = { moves: [], hidden: [], duplicates: [] }
  const warnings: string[] = []
  let names: DefinedNameEntry[] | null = null
  const tableNames = new Set(file.tableNames)
  const nextTableName = () => {
    for (let n = 1; ; n++) if (!tableNames.has(`Table${n}`)) return `Table${n}`
  }

  for (const op of ops) {
    const i = op.__index
    switch (op.op) {
      case 'set_hyperlink': {
        const sheet = name(op.sheetId)
        const { row, column } = parseAddress(op.address)
        const list = hyperlinks.get(sheet) ?? []
        list.push({ row, column, target: op.target === null ? null : linkTarget(op.target) })
        hyperlinks.set(sheet, list)
        break
      }
      case 'set_note': {
        const sheet = name(op.sheetId)
        const { row, column } = parseAddress(op.address)
        const list = (notes.get(sheet) ?? [...fileState(sheet).notes]).filter(
          (n) => !(n.row === row && n.column === column),
        )
        if (op.text !== null) list.push({ row, column, author: NOTE_AUTHOR, text: op.text })
        notes.set(sheet, list)
        break
      }
      case 'set_freeze': {
        const sheet = name(op.sheetId)
        pageSetup.set(sheet, {
          ...pageSetup.get(sheet),
          frozenRows: op.rows,
          frozenColumns: op.columns,
        })
        break
      }
      case 'set_page_setup': {
        const sheet = name(op.sheetId)
        const prior = pageSetup.get(sheet) ?? {}
        const patch: PageSetupPatch = { ...prior }
        if (op.orientation !== undefined) patch.orientation = op.orientation
        if (op.paperSize !== undefined) patch.paperSize = op.paperSize
        if (op.margins !== undefined) patch.margins = op.margins
        if (op.printGridlines !== undefined) patch.printGridlines = op.printGridlines
        if (op.printHeadings !== undefined) patch.printHeadings = op.printHeadings
        if (op.printArea !== undefined) patch.printArea = op.printArea
        if (op.printTitles !== undefined) patch.printTitles = op.printTitles
        if (op.header !== undefined) patch.header = op.header
        if (op.footer !== undefined) patch.footer = op.footer
        if (op.rowBreaks !== undefined) patch.rowBreaks = op.rowBreaks
        if (op.colBreaks !== undefined) patch.colBreaks = op.colBreaks
        // scale and fit-to-page are exclusive, as in the app: the op that sets one wins
        if (op.scale !== undefined) {
          patch.scale = op.scale
          patch.fitToPage = false
          delete patch.fitToWidth
          delete patch.fitToHeight
        } else if (op.fitToWidth !== undefined || op.fitToHeight !== undefined) {
          patch.fitToWidth = op.fitToWidth ?? prior.fitToWidth ?? 0
          patch.fitToHeight = op.fitToHeight ?? prior.fitToHeight ?? 0
          patch.fitToPage = patch.fitToWidth > 0 || patch.fitToHeight > 0
          delete patch.scale
        }
        if (Object.keys(patch).length === 0) reject(i, op.op, 'set at least one page setting')
        pageSetup.set(sheet, patch)
        break
      }
      case 'set_rows_hidden': {
        const sheet = name(op.sheetId)
        const count = op.count ?? 1
        pushHidden(hidden, sheet, {
          kind: 'set-rows-hidden',
          start: op.row - 1,
          end: op.row - 1 + count - 1,
          hidden: op.hidden,
        })
        break
      }
      case 'set_cols_hidden': {
        const sheet = name(op.sheetId)
        const start = columnIndex(op.column)
        pushHidden(hidden, sheet, {
          kind: 'set-cols-hidden',
          start,
          end: start + (op.count ?? 1) - 1,
          hidden: op.hidden,
        })
        break
      }
      case 'set_sheet_hidden':
        tabs.hidden.push({ name: name(op.sheetId), hidden: op.hidden })
        break
      case 'move_sheet':
        tabs.moves.push({ name: name(op.sheetId), position: op.position })
        break
      case 'duplicate_sheet': {
        const source = name(op.sheetId)
        const taken = new Set([
          ...file.order,
          ...after.sheets.map((s) => s.name),
          ...tabs.duplicates.map((d) => d.name),
        ])
        const copy = op.name ?? copyName(source, taken)
        if (taken.has(copy)) reject(i, op.op, `a sheet named "${copy}" already exists`)
        tabs.duplicates.push({ source, name: copy })
        break
      }
      case 'protect_sheet':
        protections.set(name(op.sheetId), op.protected)
        break
      case 'set_filter': {
        const sheet = name(op.sheetId)
        filters.set(sheet, { range: parseRange(op.range), columns: new Map(), cleared: false })
        break
      }
      case 'clear_filter': {
        const sheet = name(op.sheetId)
        const current = filters.get(sheet)?.range ?? fileState(sheet).autoFilter?.range
        if (!current) {
          warnings.push(`ops[${i}] (clear_filter): ${sheet} has no filter; nothing to clear`)
          break
        }
        filters.set(sheet, { range: current, columns: new Map(), cleared: true })
        break
      }
      case 'set_filter_criteria': {
        const sheet = name(op.sheetId)
        let draft = filters.get(sheet)
        if (draft?.cleared) {
          reject(
            i,
            op.op,
            `${sheet}'s filter was cleared earlier in this batch; run set_filter first`,
          )
        }
        if (!draft) {
          const existing = fileState(sheet).autoFilter
          if (!existing) reject(i, op.op, `${sheet} has no filter; run set_filter first`)
          if (existing!.hasCriteria) {
            reject(
              i,
              op.op,
              `${sheet}'s filter already has criteria the CLI cannot re-save; run set_filter on the range first (it clears them) or use the GenOffice app`,
            )
          }
          draft = { range: existing!.range, columns: new Map(), cleared: false }
          filters.set(sheet, draft)
        }
        const colId = columnIndex(op.column) - draft.range.startColumn
        if (colId < 0 || colId > draft.range.endColumn - draft.range.startColumn) {
          reject(i, op.op, `column ${op.column} is outside the filter range`)
        }
        if (op.values === null) draft.columns.delete(colId)
        else draft.columns.set(colId, [...op.values])
        break
      }
      case 'add_conditional_format': {
        const sheet = name(op.sheetId)
        const state = cf.get(sheet) ?? { rules: [], cleared: false }
        state.rules.push(cfWireRule(op.rule, parseRange(op.range)))
        cf.set(sheet, state)
        break
      }
      case 'clear_conditional_formats':
        cf.set(name(op.sheetId), { rules: [], cleared: true })
        break
      case 'set_data_validation': {
        const sheet = name(op.sheetId)
        const bounds = parseRange(op.range)
        const draft = dv.get(sheet) ?? { rules: [], remove: [] }
        const before = draft.rules.length
        draft.rules = draft.rules.filter((r) => !sameArea(r.ranges[0]!, bounds))
        draft.remove = draft.remove.filter((r) => !sameArea(r, bounds))
        if (op.validation === null) {
          if (fileState(sheet).dataValidations > 0) draft.remove.push(bounds)
          else if (draft.rules.length === before) {
            warnings.push(
              `ops[${i}] (set_data_validation): ${sheet}!${op.range} has no validation; nothing to remove`,
            )
          }
        } else {
          draft.rules.push({ ranges: [bounds], rule: dvWireRule(op.validation) })
        }
        dv.set(sheet, draft)
        break
      }
      case 'add_defined_name': {
        names ??= [...file.definedNames]
        const existing = names.findIndex((n) => n.name === op.name && n.sheetIndex === undefined)
        const entry = { name: op.name, formula: op.ref.replace(/^=/, '') }
        if (existing >= 0) names[existing] = entry
        else names.push(entry)
        break
      }
      case 'delete_defined_name': {
        names ??= [...file.definedNames]
        const before = names.length
        names = names.filter((n) => n.name !== op.name)
        if (names.length === before) reject(i, op.op, `no defined name "${op.name}"`)
        break
      }
      case 'add_chart':
        // built by the adapter into the snapshot; collected from the visuals below
        break
      case 'edit_chart': {
        if (!/^xl\/charts\//.test(op.chartPath)) {
          reject(
            i,
            op.op,
            'chartPath must be a chart part of the file (xl/charts/chartN.xml); `genoffice info` lists them',
          )
        }
        if (op.seriesData !== undefined) {
          reject(i, op.op, 'seriesData (repointing a series at new cells) needs the GenOffice app')
        }
        const edit: Record<string, unknown> = { chartPath: op.chartPath }
        for (const k of [
          'title',
          'chartType',
          'seriesColors',
          'legend',
          'dataLabels',
          'grouping',
          'axisTitles',
        ] as const) {
          if (op[k] !== undefined) edit[k] = op[k]
        }
        if (Object.keys(edit).length === 1) reject(i, op.op, 'set at least one chart property')
        chartEdits.push(edit as WorkbookChartEdit)
        break
      }
      case 'add_image': {
        const sheet = name(op.sheetId)
        const image = await loadImage(op.path, input.ctx, i)
        const base = parseAddress(op.anchorCell)
        // the app's frame: ~80 px per column, ~22 px per row, scaled to a ≤480 px wide picture
        const scale = Math.min(1, 480 / Math.max(1, image.width))
        const columns = Math.min(16, Math.max(2, Math.round((image.width * scale) / 80)))
        const rows = Math.min(40, Math.max(2, Math.round((image.height * scale) / 22)))
        visuals.push({
          sheetName: sheet,
          anchor: frame(base.row, base.column, rows, columns),
          image: { mediaType: image.mediaType, base64: image.base64 },
        })
        break
      }
      case 'add_shape': {
        const sheet = name(op.sheetId)
        const base = parseAddress(op.anchorCell)
        const isTextBox = op.shapeType === 'textbox'
        visuals.push({
          sheetName: sheet,
          anchor: frame(base.row, base.column, isTextBox ? 4 : 6, isTextBox ? 4 : 3),
          shape: {
            shapeType: isTextBox ? 'rect' : op.shapeType,
            fillColor: op.fillColor ?? (isTextBox ? '#FFFFFF' : '#DDEBF7'),
            ...(isTextBox
              ? { text: op.text ?? 'Text', isTextBox: true }
              : op.text === undefined
                ? {}
                : { text: op.text }),
          },
        })
        break
      }
      case 'add_table': {
        const sheet = name(op.sheetId)
        const bounds = parseRange(op.range)
        const state = sheetAfter(op.sheetId)
        const columnNames: string[] = []
        for (let c = bounds.startColumn; c <= bounds.endColumn; c++) {
          const cell = state?.cells[`${columnLabel(c)}${bounds.startRow + 1}`]
          const text =
            cell?.value === null || cell?.value === undefined ? '' : String(cell.value).trim()
          if (!text)
            reject(
              i,
              op.op,
              `header cell ${columnLabel(c)}${bounds.startRow + 1} is empty; the first row of the range holds the column names`,
            )
          if (columnNames.includes(text))
            reject(i, op.op, `duplicate column name "${text}" in the header row`)
          columnNames.push(text)
        }
        if (bounds.endRow === bounds.startRow)
          reject(i, op.op, 'the range needs at least one data row under the header')
        const tableName = op.name ?? nextTableName()
        if (tableNames.has(tableName))
          reject(i, op.op, `a table named "${tableName}" already exists`)
        tableNames.add(tableName)
        tables.push({
          sheetName: sheet,
          area: bounds,
          name: tableName,
          columnNames,
          ...(op.style === undefined ? {} : { style: op.style }),
          bandedRows: op.bandedRows ?? true,
        })
        break
      }
      case 'add_sparkline': {
        const sheet = name(op.sheetId)
        const bounds = parseRange(op.dataRange)
        const rows = bounds.endRow - bounds.startRow + 1
        if (rows > MAX_SPARKLINES_PER_OP) {
          reject(
            i,
            op.op,
            `dataRange ${op.dataRange} spans ${rows} rows; one op writes at most ${MAX_SPARKLINES_PER_OP} sparklines, split the range`,
          )
        }
        const base =
          op.targetCell === undefined
            ? { row: bounds.startRow, column: bounds.endColumn + 1 }
            : parseAddress(op.targetCell)
        const quoted = `'${sheet.replace(/'/g, "''")}'`
        sparklines.push({
          sheetName: sheet,
          type: op.type,
          ...(op.color === undefined ? {} : { color: op.color }),
          cells: Array.from({ length: rows }, (_, offset) => ({
            cell: `${columnLabel(base.column)}${base.row + offset + 1}`,
            sourceRef:
              `${quoted}!$${columnLabel(bounds.startColumn)}$${bounds.startRow + offset + 1}` +
              `:$${columnLabel(bounds.endColumn)}$${bounds.startRow + offset + 1}`,
          })),
        })
        break
      }
      case 'add_pivot': {
        const sourceName = name(op.sheetId)
        const targetId = op.targetSheetId ?? op.sheetId
        const targetName = name(targetId)
        // the writer pins both sheets' coordinates: a sheet born in this batch has none yet
        if (!file.order.includes(targetName)) {
          reject(
            i,
            op.op,
            `target sheet "${targetName}" is added in this batch; add it in a previous batch, then create the pivot`,
          )
        }
        const source = parseRange(op.sourceRange)
        const sourceCells = sheetAfter(op.sheetId)?.cells ?? {}
        const formulaCells: string[] = []
        for (let r = source.startRow; r <= source.endRow; r++) {
          for (let c = source.startColumn; c <= source.endColumn; c++) {
            const address = `${columnLabel(c)}${r + 1}`
            if (sourceCells[address]?.formula) formulaCells.push(address)
          }
        }
        let results = new Map<string, { value: PivotScalar; isError: boolean }>()
        if (formulaCells.length) {
          // the engine evaluates formulas from the file on disk, which has none of this batch's
          // cell edits; precedents can sit anywhere, so any edited cell disqualifies
          const edited = editedCells(before, after, name)
          if (edited.length) {
            reject(
              i,
              op.op,
              `the source has formulas (${formulaCells.slice(0, 5).join(', ')}) and this batch edits cells (${edited.slice(0, 5).join(', ')}); apply the edits in a previous batch, then create the pivot`,
            )
          }
          if (input.computed === undefined) {
            reject(
              i,
              op.op,
              `the source has formulas (${formulaCells.slice(0, 5).join(', ')}); needs the workbook file on disk`,
            )
          }
          results = await input.computed!(sourceName, source)
          const errors = formulaCells.filter((address) => {
            const at = parseAddress(address)
            return results.get(`${at.row}|${at.column}`)?.isError !== false
          })
          if (errors.length) {
            reject(
              i,
              op.op,
              `source formula(s) ${errors.slice(0, 5).join(', ')} evaluate to an error; fix them before creating the pivot`,
            )
          }
        }
        const grid: PivotScalar[][] = []
        for (let r = source.startRow; r <= source.endRow; r++) {
          const row: PivotScalar[] = []
          for (let c = source.startColumn; c <= source.endColumn; c++) {
            const cell = sourceCells[`${columnLabel(c)}${r + 1}`]
            if (cell?.formula) row.push(results.get(`${r}|${c}`)?.value ?? null)
            else row.push((cell?.rawValue ?? cell?.value ?? null) as PivotScalar)
          }
          grid.push(row)
        }
        let layout
        try {
          layout = buildPivotLayout(grid, op)
        } catch (err) {
          if (err instanceof PivotLayoutError) reject(i, op.op, err.message)
          throw err
        }
        const anchor = parseAddress(op.targetCell)
        const location = pivotOutputArea(anchor, layout!)
        const ref = `${columnLabel(location.startColumn)}${location.startRow + 1}:${columnLabel(location.endColumn)}${location.endRow + 1}`
        if (targetId === op.sheetId && areasOverlap(location, source)) {
          reject(i, op.op, `the pivot output ${ref} would overlap its source ${op.sourceRange}`)
        }
        const clash =
          fileState(targetName).pivots.find((p) => areasOverlap(location, p.range)) ??
          pivots.find((p) => p.sheetName === targetName && areasOverlap(location, p.location))
        if (clash) reject(i, op.op, `the pivot output ${ref} would overlap pivot "${clash.name}"`)
        const targetCells = sheetAfter(targetId)?.cells ?? {}
        const occupied = Object.entries(targetCells).filter(([address, cell]) => {
          if ((cell.value === null || cell.value === '') && !cell.formula) return false
          const at = parseAddress(address)
          return (
            at.row >= location.startRow &&
            at.row <= location.endRow &&
            at.column >= location.startColumn &&
            at.column <= location.endColumn
          )
        })
        if (occupied.length) {
          reject(
            i,
            op.op,
            `the pivot output ${ref} would overwrite ${occupied.length} non-empty cell(s) (${occupied
              .slice(0, 3)
              .map(([a]) => a)
              .join(', ')}); pick an empty targetCell or a targetSheet`,
          )
        }
        const taken = new Set(
          [...file.sheets.values()].flatMap((s) => s.pivots.map((p) => p.name.toLowerCase())),
        )
        for (const p of pivots) taken.add(p.name.toLowerCase())
        let pivotName = op.name
        if (pivotName === undefined) {
          for (let n = 1; ; n++) {
            if (!taken.has(`pivottable${n}`)) {
              pivotName = `PivotTable${n}`
              break
            }
          }
        } else if (taken.has(pivotName.toLowerCase())) {
          reject(i, op.op, `a pivot named "${pivotName}" already exists`)
        }
        pivots.push({
          sheetName: targetName,
          sourceSheetName: sourceName,
          sourceArea: source,
          location,
          name: pivotName!,
          ...layout!.definition,
        })
        const formats = new Map(layout!.numberFormats.map((f) => [f.columnOffset, f.format]))
        layout!.matrix.forEach((line, r) => {
          line.forEach((value, c) => {
            if (value === null || value === '') return
            // value columns' numFmt covers the data rows only (no header, no grand total)
            const format = r > 0 && r < layout!.height - 1 ? formats.get(c) : undefined
            baked.push({
              sheetName: targetName,
              row: anchor.row + r,
              column: anchor.column + c,
              value,
              ...(format === undefined ? {} : { numberFormat: format }),
            })
          })
        })
        break
      }
      default:
        break
    }
  }

  for (const sheet of after.sheets) {
    for (const visual of sheet.visuals ?? []) {
      if (!visual.id.startsWith('demo-chart-')) continue
      const chart = chartAddFromVisual(visual.chart)
      if (!chart) continue
      visuals.push({ sheetName: name(sheet.id), anchor: visual.anchor, chart })
    }
  }

  const filterStates: SheetFilterState[] = [...filters].map(([sheetName, draft]) => {
    const state = sheetAfter([...namesById].find(([, n]) => n === sheetName)?.[0] ?? '')
    const hiddenRows = draft.cleared ? [] : filteredRows(draft, state)
    return {
      sheetName,
      filter: draft.cleared
        ? null
        : {
            range: draft.range,
            columns: [...draft.columns].map(([colId, values]) => ({ colId, values })),
          },
      hiddenRows,
      visibilityRange: draft.range,
    }
  })

  return {
    hyperlinkEdits: [...hyperlinks].map(([sheetName, edits]) => ({ sheetName, edits })),
    noteStates: [...notes].map(([sheetName, list]) => ({ sheetName, notes: list })),
    pageSetupStates: [...pageSetup].map(([sheetName, s]) => ({ sheetName, ...s })),
    filterStates,
    cfStates: [...cf].map(([sheetName, s]) => ({ sheetName, rules: s.rules, append: !s.cleared })),
    dvStates: [...dv].map(([sheetName, d]) => ({
      sheetName,
      rules: d.rules,
      append: fileState(sheetName).dataValidations > 0,
      remove: d.remove,
    })),
    definedNamesState: names === null ? null : { names, preserveNames: [] },
    visualAdditions: visuals,
    sheetProtections: [...protections].map(([sheetName, p]) => ({ sheetName, protected: p })),
    tableAdditions: tables,
    chartEdits,
    hiddenOps: [...hidden].map(([sheetName, ops]) => ({ sheetName, ops })),
    tabs,
    pivotAdditions: pivots,
    bakedCells: baked,
    sparklineAdditions: sparklines,
    warnings,
  }
}

/** `Sheet!A1` for every cell whose value or formula differs between the snapshots */
export function editedCells(
  before: WorkbookSnapshot,
  after: WorkbookSnapshot,
  name: (sheetId: string) => string,
): string[] {
  const out: string[] = []
  for (const sheet of after.sheets) {
    const was = before.sheets.find((s) => s.id === sheet.id)?.cells ?? {}
    for (const address of new Set([...Object.keys(was), ...Object.keys(sheet.cells)])) {
      const a = was[address]
      const b = sheet.cells[address]
      if ((a?.value ?? null) !== (b?.value ?? null) || a?.formula !== b?.formula)
        out.push(`${name(sheet.id)}!${address}`)
    }
  }
  return out
}

function pushHidden(
  map: Map<string, SheetStructuralOps['ops'][number][]>,
  sheet: string,
  op: SheetStructuralOps['ops'][number],
): void {
  map.set(sheet, [...(map.get(sheet) ?? []), op])
}

function frame(row: number, column: number, rows: number, columns: number): DrawingAnchor {
  return {
    fromRow: row,
    fromColumn: column,
    fromRowOffset: 0,
    fromColumnOffset: 0,
    toRow: row + rows,
    toColumn: column + columns,
    toRowOffset: 0,
    toColumnOffset: 0,
  }
}

function sameArea(a: RangeBounds, b: RangeBounds): boolean {
  return (
    a.startRow === b.startRow &&
    a.endRow === b.endRow &&
    a.startColumn === b.startColumn &&
    a.endColumn === b.endColumn
  )
}

/** Internal references ("Sheet1!A1") become '#'-anchors; anything with a scheme is an external URL. */
function linkTarget(target: string): string {
  if (/^[a-z][a-z0-9+.-]*:/i.test(target) || target.startsWith('#')) return target
  return target.includes('!') ? `#${target}` : target
}

function copyName(source: string, taken: Set<string>): string {
  for (let n = 2; ; n++) {
    const suffix = ` (${n})`
    const candidate = `${source.slice(0, 31 - suffix.length)}${suffix}`
    if (!taken.has(candidate)) return candidate
  }
}

/** Rows of the filter span whose text is not among a column's checked values; like Excel's re-filter, the span's earlier hidden state is recomputed. */
function filteredRows(draft: FilterDraft, sheet: WorksheetState | undefined): number[] {
  const out = new Set<number>()
  for (const [colId, values] of draft.columns) {
    const wanted = new Set(values.map((v) => matchableCellText(v)))
    const column = draft.range.startColumn + colId
    for (let r = draft.range.startRow + 1; r <= draft.range.endRow; r++) {
      const cell = sheet?.cells[`${columnLabel(column)}${r + 1}`]
      const text = matchableCellText((cell?.value ?? null) as string | number | boolean | null)
      if (!wanted.has(text)) out.add(r)
    }
  }
  return [...out].sort((a, b) => a - b)
}

type CfRule = Extract<WorkbookOperation, { op: 'add_conditional_format' }>['rule']

/** DSL rule → the Univer rule model the gateway serializes (what the editor would have built). */
export function cfWireRule(rule: CfRule, bounds: RangeBounds): CfWireRule {
  const base = { ranges: [bounds], stopIfTrue: false }
  if (rule.kind === 'colorScale') {
    const stops = [
      rule.minColor,
      ...(rule.midColor === undefined ? [] : [rule.midColor]),
      rule.maxColor,
    ]
    return {
      ...base,
      rule: {
        type: 'colorScale',
        config: stops.map((color, index) => ({
          index,
          color,
          value:
            index === 0
              ? { type: 'min' }
              : index === stops.length - 1
                ? { type: 'max' }
                : { type: 'percentile', value: 50 },
        })),
      },
    }
  }
  if (rule.kind === 'dataBar') {
    return {
      ...base,
      rule: {
        type: 'dataBar',
        config: {
          min: { type: 'min' },
          max: { type: 'max' },
          positiveColor: rule.color ?? '#638EC6',
          nativeColor: '#FF555A',
        },
        isShowValue: true,
      },
    }
  }
  const style: Record<string, unknown> = {}
  if (rule.format.fillColor !== undefined) style.bg = { rgb: rule.format.fillColor }
  if (rule.format.fontColor !== undefined) style.cl = { rgb: rule.format.fontColor }
  if (rule.format.bold) style.bl = 1
  if (rule.format.italic) style.it = 1
  const highlight = (fields: Record<string, unknown>) => ({
    ...base,
    rule: { type: 'highlightCell', style, ...fields },
  })
  switch (rule.kind) {
    case 'number':
      return highlight({
        subType: 'number',
        operator: rule.operator,
        value:
          rule.operator === 'between' || rule.operator === 'notBetween'
            ? [rule.value, rule.value2 ?? rule.value]
            : rule.value,
      })
    case 'text':
      return highlight({
        subType: 'text',
        operator: {
          contains: 'containsText',
          notContains: 'notContainsText',
          beginsWith: 'beginsWith',
          endsWith: 'endsWith',
        }[rule.operator],
        value: rule.text,
      })
    case 'blank':
      return highlight({
        subType: 'text',
        operator: rule.blank ? 'containsBlanks' : 'notContainsBlanks',
      })
    case 'duplicate':
      return highlight({ subType: rule.unique ? 'uniqueValues' : 'duplicateValues' })
    case 'top10':
      return highlight({
        subType: 'rank',
        value: rule.rank,
        isPercent: rule.percent === true,
        isBottom: rule.bottom === true,
      })
    case 'formula':
      return highlight({ subType: 'formula', value: rule.formula })
  }
}

type DvRule = NonNullable<Extract<WorkbookOperation, { op: 'set_data_validation' }>['validation']>

export function dvWireRule(rule: DvRule): DvWireRule['rule'] {
  const common = { allowBlank: true, showErrorMessage: true }
  switch (rule.kind) {
    case 'list':
      return { ...common, type: 'list', formula1: rule.values.join(','), showDropDown: true }
    case 'listRef':
      return { ...common, type: 'list', formula1: `=${rule.range}`, showDropDown: true }
    case 'numberBetween':
      return {
        ...common,
        type: 'decimal',
        operator: 'between',
        formula1: String(rule.min),
        formula2: String(rule.max),
      }
    case 'dateBetween':
      return {
        ...common,
        type: 'date',
        operator: 'between',
        formula1: rule.start,
        formula2: rule.end,
      }
    case 'checkbox':
      return { ...common, type: 'checkbox' }
    case 'formula':
      return { ...common, type: 'custom', formula1: rule.formula }
  }
}

/** The chart the in-memory workbook built for add_chart → the gateway's chart part input (the app's toSaveVisualAdds). */
export function chartAddFromVisual(chart: ChartVisualState): ChartAdd | null {
  const types = chart.chartTypes
  const chartType: ChartAdd['chartType'] | null =
    types.includes('barChart') && types.includes('lineChart')
      ? 'combo'
      : types.includes('pieChart')
        ? 'pie'
        : types.includes('doughnutChart')
          ? 'doughnut'
          : types.includes('scatterChart')
            ? 'scatter'
            : types.includes('radarChart')
              ? 'radar'
              : types.includes('lineChart')
                ? 'line'
                : types.includes('areaChart')
                  ? 'area'
                  : types.includes('barChart')
                    ? chart.barDirection === 'bar'
                      ? 'bar'
                      : 'column'
                    : null
  if (chartType === null) return null
  const axisTitles = {
    ...(chart.axisTitles?.category ? { category: chart.axisTitles.category } : {}),
    ...(chart.axisTitles?.value ? { value: chart.axisTitles.value } : {}),
  }
  return {
    chartType,
    title: chart.title,
    series: chart.series.map((s) => ({
      name: s.name,
      categories: s.categories,
      values: s.values,
      ...(s.valuesRef === undefined ? {} : { valuesRef: s.valuesRef }),
      ...(s.categoriesRef === undefined ? {} : { categoriesRef: s.categoriesRef }),
      ...(s.color === undefined ? {} : { color: s.color }),
    })),
    ...(chart.legend === undefined ? {} : { legend: chart.legend }),
    ...(chart.dataLabels === undefined
      ? {}
      : {
          dataLabels:
            chart.dataLabels === 'category-value-percent' ? 'category-percent' : chart.dataLabels,
        }),
    ...(chart.dataLabelPosition === undefined
      ? {}
      : { dataLabelPosition: chart.dataLabelPosition }),
    ...(Object.keys(axisTitles).length === 0 ? {} : { axisTitles }),
    ...(chart.grouping === undefined ? {} : { grouping: chart.grouping }),
    ...(chart.gridlines === undefined ? {} : { gridlines: chart.gridlines }),
    ...(chart.valueAxis === undefined ? {} : { valueAxis: chart.valueAxis }),
    ...(chart.gapWidthPct === undefined ? {} : { gapWidthPct: Math.round(chart.gapWidthPct) }),
    ...(chart.holeSizePct === undefined ? {} : { holeSizePct: Math.round(chart.holeSizePct) }),
  }
}

interface LoadedImage {
  mediaType: 'image/png' | 'image/jpeg' | 'image/gif'
  base64: string
  width: number
  height: number
}

async function loadImage(
  source: string,
  ctx: PathContext | undefined,
  index: number,
): Promise<LoadedImage> {
  const fail = (message: string): never => reject_(index, message)
  let bytes: Uint8Array
  if (/^https?:\/\//i.test(source)) {
    const resp = await fetchRemoteImage(source)
    if (!resp || !resp.ok) fail(`could not download ${source}`)
    if (Number(resp!.headers.get('content-length') ?? 0) > MAX_IMAGE_BYTES) {
      fail(`image larger than 20 MB: ${source}`)
    }
    bytes = new Uint8Array(await resp!.arrayBuffer())
    if (bytes.length > MAX_IMAGE_BYTES) fail(`image larger than 20 MB: ${source}`)
  } else {
    const expanded = source.startsWith('~/') ? resolve(homedir(), source.slice(2)) : source
    const path = isAbsolute(expanded) ? expanded : resolve(ctx?.cwd ?? process.cwd(), expanded)
    if (!existsSync(path) || !statSync(path).isFile()) fail(`image not found: ${source}`)
    if (ctx) assertAllowed(path, ctx.env, 'read')
    if (statSync(path).size > MAX_IMAGE_BYTES) fail(`image larger than 20 MB: ${source}`)
    bytes = new Uint8Array(readFileSync(path))
  }
  const size = imageSize(bytes)
  if (
    !size ||
    (size.mime !== 'image/png' && size.mime !== 'image/jpeg' && size.mime !== 'image/gif')
  ) {
    fail(`${source} is not a PNG, JPEG or GIF image`)
  }
  return {
    mediaType: size!.mime as LoadedImage['mediaType'],
    base64: Buffer.from(bytes).toString('base64'),
    width: size!.width,
    height: size!.height,
  }
}

function reject_(index: number, message: string): never {
  throw new CliError(
    EXIT.usage,
    `ops[${index}] (add_image) rejected: ${message}`,
    { failures: [classifyOpError(index, 'add_image', message)] },
    { reason: 'op_rejected' },
  )
}
