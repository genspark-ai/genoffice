/**
 * Excel's Paste Special dialog. Univer already ships the plain paste
 * variants (values / formulas / formats / column widths / all except
 * borders) as commands that read the system clipboard; the dialog delegates
 * to them whenever no extra switch is on. Operation, Skip blanks, Transpose,
 * Paste Link and the "+ number formats" variants are ours: the copy source is
 * resolved from Univer's copy cache (or the clipboard's plain text when the
 * copy came from another app), cells are transformed here, and one
 * SetRangeValuesCommand writes the block, so it undoes in one step and
 * reaches the edit journal like any paste.
 *
 * Limits: Operation works on values (a formula in the source or target
 * contributes its cached value; Excel would wrap the formulas). A transposed
 * paste writes a formula cell's value too, since relative references would
 * need rotating, not shifting. Merged cells in the source are pasted as
 * their anchor cell. A block past the sheet edge inserts the missing
 * rows/columns first, so that paste undoes in two steps.
 */
import {
  CellValueType,
  Direction,
  ICommandService,
  IUniverInstanceService,
  type ICellData,
  type IStyleData,
  type Nullable,
  type Workbook,
  type Worksheet,
} from '@univerjs/core'
import { LexerTreeBuilder } from '@univerjs/engine-formula'
import {
  InsertColCommand,
  InsertRowCommand,
  SelectionMoveType,
  SetRangeValuesCommand,
  SetSelectionsOperation,
} from '@univerjs/sheets'
import { COPY_TYPE, ISheetClipboardService } from '@univerjs/sheets-ui'

import { formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import { sharedFormulaResolverFor } from './shared-formula-journal'
import type { UniverRuntime } from './univer-state'

export const PASTE_SPECIAL_OPEN_SHELL_COMMAND = 'paste-special-open'

export const PASTE_SPECIAL_TYPES = [
  'all',
  'formulas',
  'values',
  'formats',
  'comments',
  'validation',
  'all-source-theme',
  'all-except-borders',
  'col-widths',
  'formulas-numfmt',
  'values-numfmt',
  'all-merge-cf',
] as const
export type PasteSpecialType = (typeof PASTE_SPECIAL_TYPES)[number]

export const PASTE_SPECIAL_OPERATIONS = ['none', 'add', 'subtract', 'multiply', 'divide'] as const
export type PasteSpecialOperation = (typeof PASTE_SPECIAL_OPERATIONS)[number]

export interface PasteSpecialOptions {
  readonly type: PasteSpecialType
  readonly operation: PasteSpecialOperation
  readonly skipBlanks: boolean
  readonly transpose: boolean
  readonly link: boolean
}

/// Source cell as the transform sees it: `v`/`t` are the displayed (computed)
/// value, `f` the materialized formula text, `s` the composed style object.
export type SourceCell = Pick<ICellData, 'v' | 't' | 'f' | 'p' | 's'> | null

export type CellGrid<T> = readonly (readonly T[])[]

export type PasteSpecialSource =
  | {
      readonly kind: 'internal'
      readonly unitId: string
      readonly subUnitId: string
      readonly rows: readonly number[]
      readonly cols: readonly number[]
      readonly cut: boolean
    }
  | { readonly kind: 'text'; readonly rows: CellGrid<string> }
  | { readonly kind: 'none' }

/// Paste types the dialog disables (Univer has no note/validation-only paste).
export const UNSUPPORTED_TYPES: ReadonlySet<PasteSpecialType> = new Set(['comments', 'validation'])

const UNIVER_DELEGATE: Readonly<Record<PasteSpecialType, string | null>> = {
  all: 'univer.command.paste',
  'all-source-theme': 'univer.command.paste',
  'all-merge-cf': 'univer.command.paste',
  formulas: 'sheet.command.paste-formula',
  values: 'sheet.command.paste-value',
  formats: 'sheet.command.paste-format',
  'all-except-borders': 'sheet.command.paste-besides-border',
  'col-widths': 'sheet.command.paste-col-width',
  comments: null,
  validation: null,
  'formulas-numfmt': null,
  'values-numfmt': null,
}

const VALUE_TYPES: ReadonlySet<PasteSpecialType> = new Set([
  'all',
  'all-source-theme',
  'all-merge-cf',
  'all-except-borders',
  'formulas',
  'values',
  'formulas-numfmt',
  'values-numfmt',
])

export function pasteTypeCarriesValues(type: PasteSpecialType): boolean {
  return VALUE_TYPES.has(type)
}

/// True when the paste needs the cell transform instead of Univer's command.
export function needsOwnLayer(options: PasteSpecialOptions): boolean {
  if (options.link) return true
  // Column widths are not cell data: the switches have nothing to act on.
  if (options.type === 'col-widths') return false
  return (
    options.transpose ||
    options.skipBlanks ||
    options.operation !== 'none' ||
    UNIVER_DELEGATE[options.type] === null
  )
}

export interface PasteSpecialAvailability {
  readonly types: ReadonlySet<PasteSpecialType>
  readonly operation: boolean
  readonly skipBlanks: boolean
  readonly transpose: boolean
  readonly link: boolean
}

/// What the dialog may offer for a given copy source: a cut only pastes
/// whole, outside text has no formulas/styles/source cells to link to.
export function pasteSpecialAvailability(source: PasteSpecialSource): PasteSpecialAvailability {
  if (source.kind === 'none') {
    return { types: new Set(), operation: false, skipBlanks: false, transpose: false, link: false }
  }
  if (source.kind === 'internal' && source.cut) {
    return {
      types: new Set(['all']),
      operation: false,
      skipBlanks: false,
      transpose: false,
      link: false,
    }
  }
  if (source.kind === 'text') {
    return {
      types: new Set(['all', 'values']),
      operation: false,
      skipBlanks: false,
      transpose: true,
      link: false,
    }
  }
  return {
    types: new Set(PASTE_SPECIAL_TYPES.filter((type) => !UNSUPPORTED_TYPES.has(type))),
    operation: true,
    skipBlanks: true,
    transpose: true,
    link: true,
  }
}

export function transposeGrid<T>(grid: CellGrid<T>): T[][] {
  const rows = grid.length
  const cols = rows === 0 ? 0 : Math.max(...grid.map((row) => row.length))
  const out: T[][] = []
  for (let c = 0; c < cols; c++) {
    const line: T[] = []
    for (let r = 0; r < rows; r++) line.push(grid[r]![c] as T)
    out.push(line)
  }
  return out
}

export function isBlankSourceCell(cell: SourceCell): boolean {
  if (!cell) return true
  const hasValue = cell.v !== null && cell.v !== undefined && cell.v !== ''
  return !hasValue && !cell.f && !cell.p
}

function numericValue(cell: SourceCell | ICellData | null | undefined): number | null {
  if (!cell || cell.v === null || cell.v === undefined || cell.v === '') return null
  if (cell.t === CellValueType.BOOLEAN) return null
  if (typeof cell.v === 'number') return cell.v
  if (typeof cell.v === 'boolean') return null
  if (cell.t === CellValueType.STRING || cell.t === CellValueType.FORCE_STRING) return null
  const n = Number(cell.v)
  return Number.isFinite(n) ? n : null
}

/**
 * Excel's Operation on one cell: the pasted number combines with the number
 * already in the target. Text or boolean on either side leaves the target
 * alone; a blank target counts as 0, a blank source changes nothing.
 * Returns null when the target must stay as it is.
 */
export function applyOperation(
  target: Nullable<ICellData>,
  source: SourceCell,
  operation: PasteSpecialOperation,
): ICellData | null {
  const s = numericValue(source)
  if (s === null) return null
  const hasTarget = !!target && target.v !== null && target.v !== undefined && target.v !== ''
  const tv = hasTarget ? numericValue(target) : 0
  if (tv === null) return null
  switch (operation) {
    case 'add':
      return { v: tv + s, t: CellValueType.NUMBER, f: null, si: null }
    case 'subtract':
      return { v: tv - s, t: CellValueType.NUMBER, f: null, si: null }
    case 'multiply':
      return { v: tv * s, t: CellValueType.NUMBER, f: null, si: null }
    case 'divide':
      return s === 0
        ? { v: '#DIV/0!', t: CellValueType.STRING, f: null, si: null }
        : { v: tv / s, t: CellValueType.NUMBER, f: null, si: null }
    case 'none':
      return null
  }
}

/// `=A1` on the same sheet, `='Q1 data'!A1` across sheets (quoted like Excel).
export function linkFormula(sheetName: string | null, row: number, column: number): string {
  const address = formatAddress(row, column)
  if (sheetName === null) return `=${address}`
  const needsQuotes = !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName)
  const quoted = needsQuotes ? `'${sheetName.replace(/'/g, "''")}'` : sheetName
  return `=${quoted}!${address}`
}

export type RelocateFormula = (formula: string, rowDelta: number, columnDelta: number) => string

function withoutBorders(style: Nullable<IStyleData | string>): Nullable<IStyleData | string> {
  if (!style || typeof style === 'string') return style
  const { bd: _bd, ...rest } = style
  return rest
}

function numfmtOnly(style: Nullable<IStyleData | string>): IStyleData | undefined {
  if (!style || typeof style === 'string' || !style.n) return undefined
  return { n: style.n }
}

export interface TransformCellInput {
  readonly source: SourceCell
  readonly target: Nullable<ICellData>
  /// Offset from the source cell to the target cell (for relative references).
  readonly rowDelta: number
  readonly columnDelta: number
}

/**
 * The cell data to write for one target cell, or null to leave it untouched.
 * Pure: styles arrive composed, formulas as text, values as displayed.
 */
export function transformCell(
  input: TransformCellInput,
  options: PasteSpecialOptions,
  relocate: RelocateFormula,
): ICellData | null {
  const { source, target } = input
  if (options.skipBlanks && isBlankSourceCell(source)) return null
  if (options.operation !== 'none' && pasteTypeCarriesValues(options.type)) {
    const combined = applyOperation(target, source, options.operation)
    if (combined === null) return null
    if (options.type === 'values-numfmt' || options.type === 'formulas-numfmt') {
      const n = numfmtOnly(source?.s)
      return n ? { ...combined, s: n } : combined
    }
    if (options.type.startsWith('all')) {
      const s = options.type === 'all-except-borders' ? withoutBorders(source?.s) : source?.s
      return s ? { ...combined, s } : combined
    }
    return combined
  }
  const formula =
    source?.f && typeof source.f === 'string' && !options.transpose
      ? relocate(source.f, input.rowDelta, input.columnDelta)
      : null
  const value: ICellData = formula
    ? { f: formula, si: null, v: null, p: null }
    : { v: source?.v ?? null, t: source?.t ?? null, f: null, si: null, p: source?.p ?? null }
  const plainValue: ICellData = {
    v: source?.v ?? null,
    t: source?.t ?? null,
    f: null,
    si: null,
    p: null,
  }
  switch (options.type) {
    case 'all':
    case 'all-source-theme':
    case 'all-merge-cf':
      return source?.s ? { ...value, s: source.s } : value
    case 'all-except-borders': {
      const s = withoutBorders(source?.s)
      return s ? { ...value, s } : value
    }
    case 'formulas':
      return value
    case 'values':
      return plainValue
    case 'formats':
      return source?.s ? { s: source.s } : null
    case 'formulas-numfmt': {
      const n = numfmtOnly(source?.s)
      return n ? { ...value, s: n } : value
    }
    case 'values-numfmt': {
      const n = numfmtOnly(source?.s)
      return n ? { ...plainValue, s: n } : plainValue
    }
    case 'col-widths':
    case 'comments':
    case 'validation':
      return null
  }
}

export interface SourceGridCell {
  readonly cell: SourceCell
  /// Absolute source coordinates (null for outside text).
  readonly row: number | null
  readonly column: number | null
}

/**
 * The target-shaped grid of writes for a whole paste: transposes first, then
 * tiles when the selection is an exact multiple of the (transposed) source,
 * like Excel's repeat-paste. `null` entries leave the target cell alone.
 */
export function buildPasteGrid(
  source: CellGrid<SourceGridCell>,
  options: PasteSpecialOptions,
  target: {
    readonly startRow: number
    readonly startColumn: number
    readonly rows: number
    readonly columns: number
    readonly cellAt: (row: number, column: number) => Nullable<ICellData>
  },
  relocate: RelocateFormula,
  linkSheetName: string | null,
): { rows: number; columns: number; cells: (ICellData | null)[][] } {
  const grid = options.transpose && !options.link ? transposeGrid(source) : source
  const srcRows = grid.length
  const srcCols = srcRows === 0 ? 0 : Math.max(...grid.map((row) => row.length))
  if (srcRows === 0 || srcCols === 0) return { rows: 0, columns: 0, cells: [] }
  const tile =
    target.rows % srcRows === 0 &&
    target.columns % srcCols === 0 &&
    (target.rows > srcRows || target.columns > srcCols)
  const rows = tile ? target.rows : srcRows
  const columns = tile ? target.columns : srcCols
  const cells: (ICellData | null)[][] = []
  for (let r = 0; r < rows; r++) {
    const line: (ICellData | null)[] = []
    for (let c = 0; c < columns; c++) {
      const entry = grid[r % srcRows]?.[c % srcCols] ?? null
      const targetRow = target.startRow + r
      const targetColumn = target.startColumn + c
      if (!entry) {
        line.push(null)
        continue
      }
      if (options.link) {
        if (entry.row === null || entry.column === null) {
          line.push(null)
          continue
        }
        if (options.skipBlanks && isBlankSourceCell(entry.cell)) {
          line.push(null)
          continue
        }
        line.push({ f: linkFormula(linkSheetName, entry.row, entry.column), si: null, v: null })
        continue
      }
      line.push(
        transformCell(
          {
            source: entry.cell,
            target: target.cellAt(targetRow, targetColumn),
            rowDelta: entry.row === null ? 0 : targetRow - entry.row,
            columnDelta: entry.column === null ? 0 : targetColumn - entry.column,
          },
          options,
          relocate,
        ),
      )
    }
    cells.push(line)
  }
  return { rows, columns, cells }
}

/// Clipboard plain text → grid; quoted fields follow the TSV writer's rules.
export function parseClipboardText(text: string): string[][] {
  const normalized = text.replace(/\r\n|\r/g, '\n').replace(/\n$/, '')
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let quoted = false
  for (let i = 0; i < normalized.length; i++) {
    const ch = normalized[i]!
    if (quoted) {
      if (ch === '"') {
        if (normalized[i + 1] === '"') {
          field += '"'
          i++
        } else quoted = false
      } else field += ch
      continue
    }
    if (ch === '"' && field === '') quoted = true
    else if (ch === '\t') {
      row.push(field)
      field = ''
    } else if (ch === '\n') {
      row.push(field)
      rows.push(row)
      row = []
      field = ''
    } else field += ch
  }
  row.push(field)
  rows.push(row)
  return rows
}

export function textCell(text: string): SourceCell {
  if (text === '') return null
  if (/^(TRUE|FALSE)$/i.test(text)) {
    return { v: text.toUpperCase() === 'TRUE' ? 1 : 0, t: CellValueType.BOOLEAN }
  }
  if (/^-?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(text)) {
    const n = Number(text)
    if (Number.isFinite(n)) return { v: n, t: CellValueType.NUMBER }
  }
  return { v: text, t: CellValueType.STRING }
}

interface CopyCache {
  get(id: string):
    | {
        unitId: string
        subUnitId: string
        range: { rows: number[]; cols: number[] }
        copyType: unknown
      }
    | undefined
  getLastCopyId(): string | null
}

function copyCache(runtime: UniverRuntime): CopyCache | undefined {
  // Univer keeps the cache private on its clipboard service singleton.
  return (
    runtime.univer.__getInjector().get(ISheetClipboardService) as unknown as {
      _copyContentCache?: CopyCache
    }
  )._copyContentCache
}

function internalSource(
  cache: CopyCache | undefined,
  copyId: string | null,
): PasteSpecialSource | null {
  const entry = copyId ? cache?.get(copyId) : undefined
  if (!entry || !copyId) return null
  return {
    kind: 'internal',
    unitId: entry.unitId,
    subUnitId: entry.subUnitId,
    rows: entry.range.rows,
    cols: entry.range.cols,
    cut: entry.copyType === COPY_TYPE.CUT,
  }
}

/// The copy the dialog will paste: Univer's cache when the clipboard still
/// carries its copy id, the clipboard text when another app wrote it, else
/// the last in-app copy (the Clipboard API may be unreadable or hold only
/// an image while Univer's cache still has the copy).
export async function readPasteSpecialSource(runtime: UniverRuntime): Promise<PasteSpecialSource> {
  const cache = copyCache(runtime)
  const lastCopy = internalSource(cache, cache?.getLastCopyId() ?? null)
  let html = ''
  let text = ''
  try {
    const items = await navigator.clipboard.read()
    const item = items[0]
    if (item) {
      if (item.types.includes('text/html')) html = await (await item.getType('text/html')).text()
      if (item.types.includes('text/plain')) text = await (await item.getType('text/plain')).text()
    }
  } catch {
    return lastCopy ?? { kind: 'none' }
  }
  const copyId = /data-copy-id="([^\s"]+)"/.exec(html)?.[1] ?? null
  const internal = internalSource(cache, copyId)
  if (internal) return internal
  if (text.trim() !== '') return { kind: 'text', rows: parseClipboardText(text) }
  return lastCopy ?? { kind: 'none' }
}

function readSourceGrid(
  runtime: UniverRuntime,
  source: Extract<PasteSpecialSource, { kind: 'internal' }>,
  workbook: Workbook,
): CellGrid<SourceGridCell> | null {
  const worksheet = workbook.getSheetBySheetId(source.subUnitId)
  if (!worksheet) return null
  const resolveShared = sharedFormulaResolverFor(runtime, source.subUnitId)
  return source.rows.map((row) =>
    source.cols.map((column) => {
      const raw = worksheet.getCellRaw(row, column)
      const shown = worksheet.getCell(row, column)
      const f = raw?.f ?? (raw?.si ? resolveShared(row, column, raw.si) : null) ?? null
      const style = worksheet.getComposedCellStyleByCellData(row, column, raw)
      const cell: SourceCell =
        raw || shown
          ? {
              v: shown?.v ?? raw?.v ?? null,
              t: shown?.t ?? raw?.t ?? null,
              f,
              p: raw?.p ?? null,
              s: style && Object.keys(style).length > 0 ? style : null,
            }
          : null
      return { cell, row, column }
    }),
  )
}

export type PasteSpecialResult =
  { ok: true } | { ok: false; reason: 'no-selection' | 'out-of-bounds' | 'nothing' }

/// Runs the dialog's choice; Univer's own command for plain variants.
export async function applyPasteSpecial(
  runtime: UniverRuntime,
  source: PasteSpecialSource,
  options: PasteSpecialOptions,
): Promise<PasteSpecialResult> {
  if (source.kind === 'none') return { ok: false, reason: 'nothing' }
  if (!needsOwnLayer(options) || (source.kind === 'internal' && source.cut)) {
    const id = UNIVER_DELEGATE[options.type] ?? UNIVER_DELEGATE.all!
    await runtime.univerAPI.executeCommand(id)
    return { ok: true }
  }
  const injector = runtime.univer.__getInjector()
  const fWorkbook = runtime.univerAPI.getActiveWorkbook()
  const fSheet = fWorkbook?.getActiveSheet()
  const selection = fWorkbook?.getActiveRange()
  if (!fWorkbook || !fSheet || !selection) return { ok: false, reason: 'no-selection' }
  const targetSheet: Worksheet = fSheet.getSheet()
  let grid: CellGrid<SourceGridCell> | null
  let linkSheetName: string | null = null
  if (source.kind === 'internal') {
    const sourceWorkbook = injector.get(IUniverInstanceService).getUnit<Workbook>(source.unitId)
    if (!sourceWorkbook) return { ok: false, reason: 'nothing' }
    grid = readSourceGrid(runtime, source, sourceWorkbook)
    const sourceSheet = sourceWorkbook.getSheetBySheetId(source.subUnitId)
    linkSheetName =
      sourceSheet && source.subUnitId !== targetSheet.getSheetId() ? sourceSheet.getName() : null
  } else {
    grid = source.rows.map((row) =>
      row.map((text) => ({ cell: textCell(text), row: null, column: null })),
    )
  }
  if (!grid) return { ok: false, reason: 'nothing' }
  const lexer = injector.get(LexerTreeBuilder)
  const relocate: RelocateFormula = (formula, rowDelta, columnDelta) =>
    lexer.moveFormulaRefOffset(formula, columnDelta, rowDelta)
  const startRow = selection.getRow()
  const startColumn = selection.getColumn()
  const built = buildPasteGrid(
    grid,
    options,
    {
      startRow,
      startColumn,
      rows: selection.getHeight(),
      columns: selection.getWidth(),
      cellAt: (row, column) => targetSheet.getCell(row, column),
    },
    relocate,
    linkSheetName,
  )
  if (built.rows === 0) return { ok: false, reason: 'nothing' }
  const endRow = startRow + built.rows - 1
  const endColumn = startColumn + built.columns - 1
  const unitId = fWorkbook.getId()
  const subUnitId = targetSheet.getSheetId()
  // Like Univer's paste, grow the sheet instead of refusing the block.
  const rowCount = targetSheet.getRowCount()
  if (endRow >= rowCount) {
    await runtime.univerAPI.executeCommand(InsertRowCommand.id, {
      unitId,
      subUnitId,
      direction: Direction.DOWN,
      range: { startRow: rowCount, endRow, startColumn: 0, endColumn: 0 },
    })
  }
  const columnCount = targetSheet.getColumnCount()
  if (endColumn >= columnCount) {
    await runtime.univerAPI.executeCommand(InsertColCommand.id, {
      unitId,
      subUnitId,
      direction: Direction.RIGHT,
      range: { startRow: 0, endRow: 0, startColumn: columnCount, endColumn },
    })
  }
  if (endRow >= targetSheet.getRowCount() || endColumn >= targetSheet.getColumnCount()) {
    return { ok: false, reason: 'out-of-bounds' }
  }
  const value: Record<number, Record<number, ICellData>> = {}
  built.cells.forEach((line, r) => {
    line.forEach((cell, c) => {
      if (!cell) return
      ;(value[startRow + r] ??= {})[startColumn + c] = cell
    })
  })
  const range = { startRow, startColumn, endRow, endColumn }
  if (Object.keys(value).length > 0) {
    await runtime.univerAPI.executeCommand(SetRangeValuesCommand.id, {
      unitId,
      subUnitId,
      range,
      value,
    })
  }
  injector.get(ICommandService).syncExecuteCommand(SetSelectionsOperation.id, {
    unitId,
    subUnitId,
    type: SelectionMoveType.MOVE_END,
    selections: [{ range, primary: null }],
  })
  return { ok: true }
}
