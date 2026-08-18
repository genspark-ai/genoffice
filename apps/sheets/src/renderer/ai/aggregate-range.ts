/**
 * aggregate_range tool implementation: batched statistics that never load the
 * grid or touch the formula engine — the supported path for distinct counts
 * and frequency questions on large streamed workbooks. Extracted from App.tsx.
 */
import {
  columnLabel,
  formatAddress,
  parseAddress,
  rangeCellCount,
  type RangeBounds,
} from '../../domain/cell-address'
import type { CellScalar } from '../../domain/workbook.types'
import { isSheetRemoved, journalEntriesInRange } from '../edit-journal'
import { createRangeAggregator, type RangeAggregate } from './aggregate'
import type { WorkbookReadContext } from './workbook-readers'

export async function aggregateWorkbookRange(
  ctx: WorkbookReadContext,
  sheetIdArg: string | undefined,
  bounds: RangeBounds,
): Promise<{ ok: true; aggregate: RangeAggregate } | { ok: false; error: string }> {
  const aggregator = createRangeAggregator()
  const state = ctx.lazyWorkbookRef.current
  const workbook = ctx.univerRef.current?.univerAPI.getActiveWorkbook()
  const activeSheetId = workbook?.getActiveSheet()?.getSheetId()
  if (!state) {
    // Demo workbook: the snapshot holds every cell in memory.
    const snapshot = ctx.adapterRef.current.getSnapshot()
    const sheet = sheetIdArg
      ? snapshot.sheets.find((candidate) => candidate.id === sheetIdArg)
      : (snapshot.sheets.find((candidate) => candidate.id === activeSheetId) ?? snapshot.sheets[0])
    if (!sheet) {
      return {
        ok: false,
        error: sheetIdArg
          ? `Unknown sheet: ${sheetIdArg} (use an id from get_workbook_context)`
          : 'No workbook is open.',
      }
    }
    const worksheet = workbook?.getSheetBySheetId(sheet.id)
    let counted = 0
    for (const [address, cell] of Object.entries(sheet.cells)) {
      const position = parseAddress(address)
      if (
        position.row < bounds.startRow ||
        position.row > bounds.endRow ||
        position.column < bounds.startColumn ||
        position.column > bounds.endColumn
      ) {
        continue
      }
      let value = cell.value
      // The in-memory model stores value:null for formula cells; the computed
      // value lives in Univer's formula engine — backfill like find_cells.
      if (cell.formula && value === null && worksheet) {
        try {
          value = (worksheet.getRange(address).getValue() as CellScalar) ?? null
        } catch {
          /* fail-open: count the cell as empty */
        }
      }
      aggregator.add(value)
      counted += 1
    }
    aggregator.addEmpty(rangeCellCount(bounds) - counted)
    return { ok: true, aggregate: aggregator.finish(50) }
  }
  const sheetId = sheetIdArg ?? activeSheetId
  if (!sheetId) return { ok: false, error: 'No workbook is open.' }
  if (isSheetRemoved(state.editJournal, sheetId)) {
    return { ok: false, error: `Unknown sheet: ${sheetId}` }
  }
  const sheetMeta = state.file.sheets.find((candidate) => candidate.id === sheetId)
  if (!sheetMeta) {
    return {
      ok: false,
      error:
        'aggregate_range only works on sheets that came with the file — read sheets added this session with read_range.',
    }
  }
  if ((state.editJournal.structuralOps.get(sheetId) ?? []).length > 0) {
    return {
      ok: false,
      error:
        'aggregate_range is unavailable after row/column structural changes on this sheet in this session — save the file first, then retry.',
    }
  }
  // Cells written this session can sit below/right of the file extent.
  let extentRows = sheetMeta.rowCount
  let extentColumns = sheetMeta.columnCount
  for (const entry of state.editJournal.cells.get(sheetId)?.values() ?? []) {
    if (!entry.hasValue) continue
    extentRows = Math.max(extentRows, entry.row + 1)
    extentColumns = Math.max(extentColumns, entry.column + 1)
  }
  const clamped = {
    startRow: bounds.startRow,
    endRow: Math.min(bounds.endRow, extentRows - 1),
    startColumn: bounds.startColumn,
    endColumn: Math.min(bounds.endColumn, extentColumns - 1),
  }
  if (clamped.startRow > clamped.endRow || clamped.startColumn > clamped.endColumn) {
    return {
      ok: false,
      error: `The range is outside the sheet data extent A1:${columnLabel(extentColumns - 1)}${extentRows}.`,
    }
  }
  const worksheet = workbook?.getSheetBySheetId(sheetId)
  // Session edits win over the streamed file content.
  const journalValues = new Map<string, CellScalar>()
  for (const entry of journalEntriesInRange(state.editJournal, sheetId, clamped)) {
    if (!entry.hasValue) continue
    let value = entry.value ?? null
    // Journal formula entries store value:null; the computed result lives in
    // Univer (journal edits are always applied there) — backfill like find_cells.
    if (entry.formula && value === null && worksheet) {
      try {
        value =
          (worksheet.getRange(formatAddress(entry.row, entry.column)).getValue() as CellScalar) ??
          null
      } catch {
        /* fail-open: count the cell as empty */
      }
    }
    journalValues.set(`${entry.row}:${entry.column}`, value)
  }
  let counted = 0
  const fileEndRow = Math.min(clamped.endRow, sheetMeta.rowCount - 1)
  const fileEndColumn = Math.min(clamped.endColumn, sheetMeta.columnCount - 1)
  if (clamped.startRow <= fileEndRow && clamped.startColumn <= fileEndColumn) {
    const width = fileEndColumn - clamped.startColumn + 1
    const batchRows = Math.max(1, Math.floor(18_000 / width))
    for (let startRow = clamped.startRow; startRow <= fileEndRow; startRow += batchRows) {
      const endRow = Math.min(startRow + batchRows - 1, fileEndRow)
      let result
      try {
        result = await window.desktopApi.readWorkbookRange({
          sessionId: state.file.sessionId,
          sheetId,
          range: {
            startRow,
            endRow,
            startColumn: clamped.startColumn,
            endColumn: fileEndColumn,
          },
        })
      } catch (error: unknown) {
        return {
          ok: false,
          error: error instanceof Error ? error.message : 'Failed to read the range.',
        }
      }
      if (result.indexedThroughRow === null || result.indexedThroughRow < endRow) {
        return {
          ok: false,
          error: 'The range is still being indexed — retry after workbook indexing completes.',
        }
      }
      for (const cell of result.cells) {
        if (journalValues.has(`${cell.row}:${cell.column}`)) continue
        aggregator.add(cell.value)
        counted += 1
      }
    }
  }
  for (const value of journalValues.values()) {
    if (value !== null && value !== '') {
      aggregator.add(value)
      counted += 1
    }
  }
  aggregator.addEmpty(rangeCellCount(clamped) - counted)
  return { ok: true, aggregate: aggregator.finish(50) }
}
