import type { LazyWorkbookState, UniverWorksheet } from './univer-state'
import { inferContinuousRegion, type TableRegion } from './table-actions'
import { shiftCellArea, StructuralShiftError } from '@genoffice/xlsx-gateway/gateway/xlsx-structure'

/// Infer only from a complete model: unloaded cells look like blank boundaries.
/// Explicit selections are always authoritative, including single-row ranges.
export function resolvePivotSource(
  worksheet: UniverWorksheet,
  selection: TableRegion,
  state: LazyWorkbookState | null,
): TableRegion {
  if (
    selection.startRow !== selection.endRow ||
    selection.startColumn !== selection.endColumn ||
    !state?.flags.preloadComplete
  ) {
    return selection
  }
  const sheetId = worksheet.getSheetId()
  const contains = (range: TableRegion) =>
    selection.startRow >= range.startRow &&
    selection.startRow <= range.endRow &&
    selection.startColumn >= range.startColumn &&
    selection.startColumn <= range.endColumn

  // File tables are not necessarily installed in Univer's table service.
  // Map their bounds into the live grid and omit totals to avoid double counting.
  const ops = state.editJournal.structuralOps.get(sheetId) ?? []
  for (const table of state.file.sheets.find((sheet) => sheet.id === sheetId)?.tables ?? []) {
    try {
      // Match the save path, including rows inserted inside the table.
      const range = shiftCellArea(table.range, ops)
      if (!range || !contains(range)) continue
      if (table.headerRowCount !== 1) return selection
      const source = shiftCellArea(
        {
          ...table.range,
          endRow: table.range.endRow - (table.totalsRowCount ?? 0),
        },
        ops,
      )
      return source && source.endRow > source.startRow ? source : selection
    } catch (error) {
      if (error instanceof StructuralShiftError) return selection
      throw error
    }
  }
  for (const table of state.editJournal.tableAdds) {
    if (table.sheetId === sheetId && contains(table.area)) return table.area
  }
  return inferContinuousRegion(worksheet, selection.startRow, selection.startColumn) ?? selection
}
