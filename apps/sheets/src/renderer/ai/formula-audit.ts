/**
 * Formula auditing for the AI skill layer (trace_precedents /
 * trace_dependents). Precedents parse the target formula's references (same
 * tokenizer as the streaming closure and the ribbon trace arrows) and sample
 * their current values; dependents scan every sheet's formulas — the file
 * model via the readWorkbookFormulas sidecar channel with session edits
 * overlaid from the journal — for references covering the target cell.
 */
import { formatAddress, parseAddress } from '../../domain/cell-address'
import type { CellScalar } from '../../domain/workbook.types'
import { qualifierMatches, type StructuralOp } from '../../gateway/xlsx-structure'
import { containsUnresolvedNames, parseFormulaReferences } from '../formula-closure'
import { readSheetRangeMapped } from '../univer-sync'
import { fileRangeToScreenRange, fileToScreen, netAxisDelta } from '../view-transform'
import type {
  TraceCellSample,
  TraceDependentInfo,
  TraceDependentsOutcome,
  TracePrecedentsOutcome,
  TraceRefInfo,
} from './tools'
import type { WorkbookReadContext } from './workbook-readers'
import { ERROR_VALUE_RE } from './workbook-search'

const MAX_REFS = 30
const MAX_SAMPLE_CELLS = 12
const MAX_DEPENDENTS = 100
/** Stand-in bound for whole-row/column references before clamping/mapping. */
const AXIS_MAX = 9_999_999

interface SheetInfo {
  readonly id: string
  readonly name: string
  /** original file sheet name — differs from name after a session rename;
   * fromFile formulas still qualify refs with this name until save */
  readonly fileName?: string | undefined
}

/** Whether a formula's sheet qualifier resolves to this sheet. File formulas
 * can only reference sheets that existed in the file — a session-added sheet
 * (no fileName) must never match them, even if it reuses a renamed sheet's
 * old name. */
function qualifierHitsSheet(qualifier: string, sheet: SheetInfo, fromFile: boolean): boolean {
  if (fromFile) {
    return sheet.fileName !== undefined && qualifierMatches(qualifier, sheet.fileName)
  }
  return qualifierMatches(qualifier, sheet.name)
}

interface RectBounds {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

function listSheets(ctx: WorkbookReadContext): { sheets: SheetInfo[]; activeId: string } | null {
  const state = ctx.lazyWorkbookRef.current
  if (state) {
    const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
    const active = workbook?.getActiveSheet()
    if (!workbook || !active) return null
    return {
      sheets: workbook.getSheets().map((sheet) => ({
        id: sheet.getSheetId(),
        name: sheet.getSheetName(),
        fileName: state.file.sheets.find((meta) => meta.id === sheet.getSheetId())?.name,
      })),
      activeId: active.getSheetId(),
    }
  }
  const snapshot = ctx.adapterRef.current.getSnapshot()
  const first = snapshot.sheets[0]
  if (!first) return null
  // Demo Univer sheets reuse the snapshot ids, so the grid knows which one
  // is active; the adapter does not.
  const activeId = ctx.univerRef.current?.univerAPI
    .getActiveWorkbook()
    ?.getActiveSheet()
    ?.getSheetId()
  return {
    sheets: snapshot.sheets.map((sheet) => ({ id: sheet.id, name: sheet.name })),
    activeId: snapshot.sheets.some((sheet) => sheet.id === activeId)
      ? (activeId as string)
      : first.id,
  }
}

/** Data extent in screen coordinates (rows, columns); zeros when unknown. */
function sheetExtent(ctx: WorkbookReadContext, sheetId: string): { rows: number; columns: number } {
  const state = ctx.lazyWorkbookRef.current
  if (state) {
    const meta = state.file.sheets.find((sheet) => sheet.id === sheetId)
    if (!meta) return { rows: 0, columns: 0 }
    const ops = state.editJournal.structuralOps.get(sheetId) ?? []
    return {
      rows: Math.max(meta.rowCount + netAxisDelta(ops, 'row'), 0),
      columns: Math.max(meta.columnCount + netAxisDelta(ops, 'column'), 0),
    }
  }
  const sheet = ctx.adapterRef.current.getSnapshot().sheets.find((entry) => entry.id === sheetId)
  if (!sheet) return { rows: 0, columns: 0 }
  let rows = 0
  let columns = 0
  for (const address of Object.keys(sheet.cells)) {
    const cell = parseAddress(address)
    if (cell.row + 1 > rows) rows = cell.row + 1
    if (cell.column + 1 > columns) columns = cell.column + 1
  }
  return { rows, columns }
}

function backfillFormulaValue(
  ctx: WorkbookReadContext,
  sheetId: string,
  address: string,
  value: CellScalar,
  formula: string | undefined,
): CellScalar {
  if (!formula || value !== null) return value
  try {
    const worksheet = ctx.univerRef.current?.univerAPI
      .getActiveWorkbook()
      ?.getSheetBySheetId(sheetId)
    return (worksheet?.getRange(address).getValue() as CellScalar) ?? null
  } catch {
    return value
  }
}

/** Row-major cell samples of a rect, journal-overlaid in lazy mode; empty
 * cells are included so the model sees gaps in a precedent range. */
async function sampleCells(
  ctx: WorkbookReadContext,
  sheetId: string,
  bounds: RectBounds,
  limit: number,
): Promise<TraceCellSample[]> {
  const width = bounds.endColumn - bounds.startColumn + 1
  const rowsNeeded = Math.max(1, Math.ceil(limit / Math.max(width, 1)))
  const clamped: RectBounds = {
    ...bounds,
    endRow: Math.min(bounds.endRow, bounds.startRow + rowsNeeded - 1),
  }
  const cells = new Map<string, { value: CellScalar; formula?: string | undefined }>()
  const state = ctx.lazyWorkbookRef.current
  if (state) {
    const meta = state.file.sheets.find((sheet) => sheet.id === sheetId)
    if (meta) {
      try {
        const mapped = await readSheetRangeMapped(state, sheetId, { ...clamped }, meta)
        for (const cell of mapped?.screen.cells ?? []) {
          cells.set(`${cell.row}:${cell.column}`, { value: cell.value, formula: cell.formula })
        }
      } catch {
        /* fail-open: journal overlay below still applies */
      }
    }
    for (const entry of state.editJournal.cells.get(sheetId)?.values() ?? []) {
      if (!entry.hasValue) continue
      if (
        entry.row < clamped.startRow ||
        entry.row > clamped.endRow ||
        entry.column < clamped.startColumn ||
        entry.column > clamped.endColumn
      ) {
        continue
      }
      cells.set(`${entry.row}:${entry.column}`, { value: entry.value, formula: entry.formula })
    }
  } else {
    const sheet = ctx.adapterRef.current.getSnapshot().sheets.find((entry) => entry.id === sheetId)
    for (const [address, cell] of Object.entries(sheet?.cells ?? {})) {
      const point = parseAddress(address)
      if (
        point.row < clamped.startRow ||
        point.row > clamped.endRow ||
        point.column < clamped.startColumn ||
        point.column > clamped.endColumn
      ) {
        continue
      }
      cells.set(`${point.row}:${point.column}`, { value: cell.value, formula: cell.formula })
    }
  }
  const samples: TraceCellSample[] = []
  for (let row = clamped.startRow; row <= clamped.endRow && samples.length < limit; row += 1) {
    for (
      let column = clamped.startColumn;
      column <= clamped.endColumn && samples.length < limit;
      column += 1
    ) {
      const address = formatAddress(row, column)
      const cell = cells.get(`${row}:${column}`)
      const value = backfillFormulaValue(ctx, sheetId, address, cell?.value ?? null, cell?.formula)
      samples.push({ address, value, formula: cell?.formula })
    }
  }
  return samples
}

/** One cell with its provenance: file formulas keep their reference text in
 * FILE coordinates until save rewrites them, so callers must map refs of
 * fromFile formulas through the sheet's journaled structural ops. */
async function readAuditCell(
  ctx: WorkbookReadContext,
  sheetId: string,
  row: number,
  column: number,
): Promise<{ value: CellScalar; formula?: string | undefined; fromFile: boolean } | undefined> {
  const state = ctx.lazyWorkbookRef.current
  if (!state) {
    const sheet = ctx.adapterRef.current.getSnapshot().sheets.find((entry) => entry.id === sheetId)
    const cell = sheet?.cells[formatAddress(row, column)]
    return cell ? { value: cell.value, formula: cell.formula, fromFile: false } : undefined
  }
  for (const entry of state.editJournal.cells.get(sheetId)?.values() ?? []) {
    if (!entry.hasValue || entry.row !== row || entry.column !== column) continue
    return { value: entry.value, formula: entry.formula, fromFile: false }
  }
  const meta = state.file.sheets.find((entry) => entry.id === sheetId)
  if (!meta) return undefined
  try {
    const mapped = await readSheetRangeMapped(
      state,
      sheetId,
      { startRow: row, endRow: row, startColumn: column, endColumn: column },
      meta,
    )
    const cell = mapped?.screen.cells.find((entry) => entry.row === row && entry.column === column)
    return cell ? { value: cell.value, formula: cell.formula, fromFile: true } : undefined
  } catch {
    return undefined
  }
}

function rectLabel(sheet: SheetInfo, bounds: RectBounds, sameSheet: boolean): string {
  const prefix = sameSheet ? '' : `${sheet.name}!`
  const single =
    bounds.startRow === bounds.endRow && bounds.startColumn === bounds.endColumn
      ? formatAddress(bounds.startRow, bounds.startColumn)
      : `${formatAddress(bounds.startRow, bounds.startColumn)}:${formatAddress(bounds.endRow, bounds.endColumn)}`
  return `${prefix}${single}`
}

export async function traceWorkbookPrecedents(
  ctx: WorkbookReadContext,
  sheetIdInput: string | undefined,
  address: string,
): Promise<TracePrecedentsOutcome> {
  const resolved = listSheets(ctx)
  if (!resolved) return { refs: [], error: 'No workbook is currently open.' }
  const sheetId = sheetIdInput ?? resolved.activeId
  const sheet = resolved.sheets.find((entry) => entry.id === sheetId)
  if (!sheet) return { refs: [], error: `Unknown sheet: ${sheetIdInput}` }
  const target = parseAddress(address)
  const cell = await readAuditCell(ctx, sheet.id, target.row, target.column)
  if (!cell?.formula) {
    return { refs: [], value: cell?.value ?? null }
  }
  const state = ctx.lazyWorkbookRef.current
  const references = parseFormulaReferences(cell.formula)
  const refs: TraceRefInfo[] = []
  for (const reference of references.slice(0, MAX_REFS)) {
    const targetSheet =
      reference.qualifier === undefined
        ? sheet
        : resolved.sheets.find((entry) =>
            qualifierHitsSheet(reference.qualifier!, entry, cell.fromFile),
          )
    if (!targetSheet) {
      refs.push({
        label: `${reference.qualifier}!…`,
        cellCount: 0,
        samples: [],
        hasError: false,
        external: true,
      })
      continue
    }
    const token = reference.token
    // File formulas keep refs in file coordinates until save; map them
    // through the referenced sheet's journaled structural ops
    const refOps = cell.fromFile ? (state?.editJournal.structuralOps.get(targetSheet.id) ?? []) : []
    let bounds: RectBounds
    if (refOps.length > 0) {
      const meta = state?.file.sheets.find((entry) => entry.id === targetSheet.id)
      const fileRect: RectBounds = {
        startRow: Math.max(token.startRow ?? 0, 0),
        endRow: Math.max(token.endRow ?? Math.max((meta?.rowCount ?? 1) - 1, 0), 0),
        startColumn: Math.max(token.startColumn ?? 0, 0),
        endColumn: Math.max(token.endColumn ?? Math.max((meta?.columnCount ?? 1) - 1, 0), 0),
      }
      const mapped = fileRangeToScreenRange(refOps, fileRect)
      // Range fully deleted this session — nothing left to sample
      if (!mapped) continue
      // Whole-axis references re-open to the screen extent so lines inserted
      // on that axis stay covered
      const screenExtent = sheetExtent(ctx, targetSheet.id)
      bounds = {
        startRow: token.startRow === null ? 0 : Math.max(mapped.startRow, 0),
        endRow:
          token.endRow === null ? Math.max(screenExtent.rows - 1, 0) : Math.max(mapped.endRow, 0),
        startColumn: token.startColumn === null ? 0 : Math.max(mapped.startColumn, 0),
        endColumn:
          token.endColumn === null
            ? Math.max(screenExtent.columns - 1, 0)
            : Math.max(mapped.endColumn, 0),
      }
    } else {
      const extent = sheetExtent(ctx, targetSheet.id)
      bounds = {
        startRow: Math.max(token.startRow ?? 0, 0),
        endRow: Math.max(token.endRow ?? Math.max(extent.rows - 1, 0), 0),
        startColumn: Math.max(token.startColumn ?? 0, 0),
        endColumn: Math.max(token.endColumn ?? Math.max(extent.columns - 1, 0), 0),
      }
    }
    if (bounds.endRow < bounds.startRow || bounds.endColumn < bounds.startColumn) continue
    const cellCount =
      (bounds.endRow - bounds.startRow + 1) * (bounds.endColumn - bounds.startColumn + 1)
    const samples = await sampleCells(ctx, targetSheet.id, bounds, MAX_SAMPLE_CELLS)
    refs.push({
      label: rectLabel(targetSheet, bounds, targetSheet.id === sheet.id),
      cellCount,
      samples,
      hasError: samples.some(
        (sample) => typeof sample.value === 'string' && ERROR_VALUE_RE.test(sample.value),
      ),
    })
  }
  return {
    formula: cell.formula,
    value: backfillFormulaValue(ctx, sheet.id, address, cell.value, cell.formula),
    refs,
    truncatedRefs: references.length > MAX_REFS,
    usesNames: containsUnresolvedNames(cell.formula),
  }
}

interface SheetFormulaCell {
  readonly row: number
  readonly column: number
  readonly formula: string
  readonly value: CellScalar
  /** formula text still in file coordinates (see readAuditCell) */
  readonly fromFile: boolean
}

async function collectSheetFormulas(
  ctx: WorkbookReadContext,
  sheet: SheetInfo,
): Promise<{ cells: SheetFormulaCell[]; complete: boolean }> {
  const state = ctx.lazyWorkbookRef.current
  if (!state) {
    const snapshot = ctx.adapterRef.current
      .getSnapshot()
      .sheets.find((entry) => entry.id === sheet.id)
    const cells: SheetFormulaCell[] = []
    for (const [address, cell] of Object.entries(snapshot?.cells ?? {})) {
      if (!cell.formula) continue
      const point = parseAddress(address)
      cells.push({
        row: point.row,
        column: point.column,
        formula: cell.formula,
        value: backfillFormulaValue(ctx, sheet.id, address, cell.value, cell.formula),
        fromFile: false,
      })
    }
    return { cells, complete: true }
  }
  const cells: SheetFormulaCell[] = []
  const shadowed = new Set<string>()
  for (const entry of state.editJournal.cells.get(sheet.id)?.values() ?? []) {
    if (!entry.hasValue) continue
    shadowed.add(`${entry.row}:${entry.column}`)
    if (!entry.formula) continue
    const address = formatAddress(entry.row, entry.column)
    cells.push({
      row: entry.row,
      column: entry.column,
      formula: entry.formula,
      value: backfillFormulaValue(ctx, sheet.id, address, entry.value, entry.formula),
      fromFile: false,
    })
  }
  const meta = state.file.sheets.find((entry) => entry.id === sheet.id)
  if (!meta) return { cells, complete: true }
  let result
  try {
    result = await window.desktopApi.readWorkbookFormulas({
      sessionId: state.file.sessionId,
      sheetId: sheet.id,
    })
  } catch {
    return { cells, complete: false }
  }
  const ops = state.editJournal.structuralOps.get(sheet.id) ?? []
  for (const cell of result.cells) {
    if (!cell.formula) continue
    const row = fileToScreen(ops, 'row', cell.row)
    const column = fileToScreen(ops, 'column', cell.column)
    if (row === null || column === null) continue
    if (shadowed.has(`${row}:${column}`)) continue
    cells.push({ row, column, formula: cell.formula, value: cell.value, fromFile: true })
  }
  return { cells, complete: result.indexingComplete && !result.truncated }
}

export async function traceWorkbookDependents(
  ctx: WorkbookReadContext,
  sheetIdInput: string | undefined,
  address: string,
): Promise<TraceDependentsOutcome> {
  const resolved = listSheets(ctx)
  if (!resolved) {
    return {
      dependents: [],
      truncated: false,
      incompleteSheets: [],
      error: 'No workbook is currently open.',
    }
  }
  const sheetId = sheetIdInput ?? resolved.activeId
  const targetSheet = resolved.sheets.find((entry) => entry.id === sheetId)
  if (!targetSheet) {
    return {
      dependents: [],
      truncated: false,
      incompleteSheets: [],
      error: `Unknown sheet: ${sheetIdInput}`,
    }
  }
  const target = parseAddress(address)
  // File formulas keep refs in the target sheet's FILE coordinates until
  // save; map those refs through the target sheet's structural ops
  const targetOps = ctx.lazyWorkbookRef.current?.editJournal.structuralOps.get(targetSheet.id) ?? []
  const dependents: TraceDependentInfo[] = []
  const incompleteSheets: string[] = []
  let truncated = false
  for (const sheet of resolved.sheets) {
    const { cells, complete } = await collectSheetFormulas(ctx, sheet)
    if (!complete) incompleteSheets.push(sheet.name)
    for (const cell of cells) {
      if (sheet.id === targetSheet.id && cell.row === target.row && cell.column === target.column) {
        continue
      }
      if (
        !formulaReadsTarget(
          cell.formula,
          sheet.id === targetSheet.id,
          // undefined = file formulas cannot reference this (session-added) sheet
          cell.fromFile ? targetSheet.fileName : targetSheet.name,
          target,
          cell.fromFile ? targetOps : [],
        )
      ) {
        continue
      }
      if (dependents.length >= MAX_DEPENDENTS) {
        truncated = true
        break
      }
      dependents.push({
        sheetName: sheet.name,
        address: formatAddress(cell.row, cell.column),
        formula: cell.formula,
        value: cell.value,
      })
    }
    if (truncated) break
  }
  return { dependents, truncated, incompleteSheets }
}

function formulaReadsTarget(
  formula: string,
  onTargetSheet: boolean,
  targetSheetName: string | undefined,
  target: { row: number; column: number },
  refOps: readonly StructuralOp[],
): boolean {
  for (const reference of parseFormulaReferences(formula)) {
    const hitsTargetSheet =
      reference.qualifier === undefined
        ? onTargetSheet
        : targetSheetName !== undefined && qualifierMatches(reference.qualifier, targetSheetName)
    if (!hitsTargetSheet) continue
    const { token } = reference
    let rect: RectBounds = {
      startRow: token.startRow ?? 0,
      endRow: token.endRow ?? AXIS_MAX,
      startColumn: token.startColumn ?? 0,
      endColumn: token.endColumn ?? AXIS_MAX,
    }
    if (refOps.length > 0) {
      const mapped = fileRangeToScreenRange(refOps, rect)
      if (!mapped) continue
      // Whole-axis references stay whole-axis: an insert on the open axis
      // must not push inserted lines outside the ref (A:A covers new rows)
      rect = {
        startRow: token.startRow === null ? 0 : mapped.startRow,
        endRow: token.endRow === null ? AXIS_MAX : mapped.endRow,
        startColumn: token.startColumn === null ? 0 : mapped.startColumn,
        endColumn: token.endColumn === null ? AXIS_MAX : mapped.endColumn,
      }
    }
    if (
      rect.startRow <= target.row &&
      target.row <= rect.endRow &&
      rect.startColumn <= target.column &&
      target.column <= rect.endColumn
    ) {
      return true
    }
  }
  return false
}
