import {
  FORMULA_MODE_MAX_CELLS,
  FULL_LOAD_MAX_CELLS,
  FULL_LOAD_MAX_GRID_CELLS,
  FULL_LOAD_MAX_STORED_CELLS,
} from './app-constants'

export interface CellCountedSheet {
  readonly rowCount: number
  readonly columnCount: number
  readonly storedCellCount?: number | undefined
  readonly valueCellCount?: number | undefined
  readonly valueRowCount?: number | undefined
  readonly valueColumnCount?: number | undefined
}

export interface WorkbookCellCounts {
  /// Sum of the boxes the value cells span; a sheet without a reported
  /// extent (older binary, count budget spent) contributes its bounding box.
  readonly gridCells: number
  /// Sum of stored cells (value, formula or style); falls back to the
  /// bounding box per sheet when the sidecar did not report a count.
  readonly storedCells: number
  /// Sum of value/formula cells; falls back to storedCells per sheet.
  readonly valueCells: number
}

export function workbookCellCounts(sheets: readonly CellCountedSheet[]): WorkbookCellCounts {
  let gridCells = 0
  let storedCells = 0
  let valueCells = 0
  for (const sheet of sheets) {
    const box = sheet.rowCount * sheet.columnCount
    const stored = sheet.storedCellCount ?? box
    const hasValueBox = sheet.valueRowCount !== undefined && sheet.valueColumnCount !== undefined
    gridCells += hasValueBox ? sheet.valueRowCount * sheet.valueColumnCount : box
    storedCells += stored
    valueCells += sheet.valueCellCount ?? stored
  }
  return { gridCells, storedCells, valueCells }
}

/// Formula mode preloads the whole workbook at open, so it is bounded like
/// Full Load and additionally by the formulas the engine must hold.
export function opensInFormulaMode(counts: WorkbookCellCounts): boolean {
  return counts.storedCells <= FORMULA_MODE_MAX_CELLS && fitsFullLoad(counts)
}

/// Full Load makes every stored cell resident and the block install writes
/// a dense matrix over the value box, so all three counts bound it.
export function fitsFullLoad(counts: WorkbookCellCounts): boolean {
  return (
    counts.valueCells <= FULL_LOAD_MAX_CELLS &&
    counts.storedCells <= FULL_LOAD_MAX_STORED_CELLS &&
    counts.gridCells <= FULL_LOAD_MAX_GRID_CELLS
  )
}
