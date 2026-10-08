import {
  FORMULA_MODE_MAX_CELLS,
  FULL_LOAD_MAX_CELLS,
  FULL_LOAD_MAX_GRID_CELLS,
} from './app-constants'

export interface CellCountedSheet {
  readonly rowCount: number
  readonly columnCount: number
  readonly storedCellCount?: number | undefined
}

export interface WorkbookCellCounts {
  /// Sum of each sheet's bounding box — what viewport math scales with.
  readonly gridCells: number
  /// Sum of stored cells; falls back to gridCells per sheet when the sidecar
  /// did not report a count (older binary).
  readonly storedCells: number
}

export function workbookCellCounts(sheets: readonly CellCountedSheet[]): WorkbookCellCounts {
  let gridCells = 0
  let storedCells = 0
  for (const sheet of sheets) {
    const grid = sheet.rowCount * sheet.columnCount
    gridCells += grid
    storedCells += sheet.storedCellCount ?? grid
  }
  return { gridCells, storedCells }
}

/// Formula mode preloads the whole workbook at open, so it is bounded like
/// Full Load and additionally by the formulas the engine must hold.
export function opensInFormulaMode(counts: WorkbookCellCounts): boolean {
  return counts.storedCells <= FORMULA_MODE_MAX_CELLS && fitsFullLoad(counts)
}

/// Full Load makes every stored cell resident, and the block install still
/// writes a dense matrix over the box, so both counts bound it.
export function fitsFullLoad(counts: WorkbookCellCounts): boolean {
  return counts.storedCells <= FULL_LOAD_MAX_CELLS && counts.gridCells <= FULL_LOAD_MAX_GRID_CELLS
}
