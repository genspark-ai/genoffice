/**
 * What-If Analysis › Data Table: substitute a run of values (one or two
 * variables) through one formula cell and write each outcome into the grid,
 * Excel's Data Table semantics. Excel stores the result area as a
 * `{=TABLE(rowInput, colInput)}` array formula; that is not reproducible
 * here — Univer's engine has no TABLE function (the formula would evaluate
 * to #NAME? on every recalculation and destroy the results) and the edit
 * journal / save path carries no array-formula channel (`t="array"` + ref),
 * so results are written as plain values. That is the honest divergence:
 * unlike Excel, the table does not refresh when the model changes.
 *
 * Each substitution is Goal Seek's mutateAndAwaitRecalc pattern — a journaled
 * setValue, then wait for the recalculation it triggers (Calculate Now in
 * manual mode), then read the formula cell. The whole run sits inside one
 * undo batch, so a single ⌘Z removes every result and restores the input
 * cells; on any failure the input cells are restored before the throw.
 */
import {
  parseAddress,
  parseRange,
  type CellCoordinates,
  type RangeBounds,
} from '@genoffice/xlsx-gateway/domain/cell-address'

import { calculateNow, isManualCalculation } from './calc-options'
import { t } from './i18n/locale'
import { beginUndoBatch } from './op-executor'
import type { UniverRuntime, UniverWorksheet } from './univer-state'

/// Excel caps nothing; we do, because every result cell is a full
/// mutate-recalc-read cycle and the run must stay undoable in one step.
export const MAX_DATA_TABLE_CELLS = 10_000

/** One-variable tables oriented down a column, across a row, or two-variable. */
export type DataTableKind = 'column' | 'row' | 'grid'

export type DataTablePlanError =
  'noInputCell' | 'sameInputCell' | 'badShape' | 'inputInsideTable' | 'tooManyCells'

export interface DataTablePlan {
  readonly kind: DataTableKind
  /// The cell whose formula is re-evaluated for every substitution.
  readonly formulaCell: CellCoordinates
  /// Cells holding the values substituted into the row input cell (the row
  /// across the table's top edge). Empty on column-oriented tables.
  readonly rowValueCells: readonly CellCoordinates[]
  /// Cells holding the values substituted into the column input cell (the
  /// column down the table's left edge). Empty on row-oriented tables.
  readonly columnValueCells: readonly CellCoordinates[]
  /// The results block: resultRows × resultColumns at resultStart.
  readonly resultStart: CellCoordinates
  readonly resultRows: number
  readonly resultColumns: number
}

const inside = (bounds: RangeBounds, cell: CellCoordinates): boolean =>
  cell.row >= bounds.startRow &&
  cell.row <= bounds.endRow &&
  cell.column >= bounds.startColumn &&
  cell.column <= bounds.endColumn

const cellsDownColumn = (bounds: RangeBounds): CellCoordinates[] =>
  Array.from({ length: bounds.endRow - bounds.startRow }, (_, index) => ({
    row: bounds.startRow + 1 + index,
    column: bounds.startColumn,
  }))

const cellsAcrossRow = (bounds: RangeBounds): CellCoordinates[] =>
  Array.from({ length: bounds.endColumn - bounds.startColumn }, (_, index) => ({
    row: bounds.startRow,
    column: bounds.startColumn + 1 + index,
  }))

/**
 * Maps the selected region onto Excel's Data Table layouts. `rowInput` is the
 * cell the values across the top row are substituted into; `columnInput` the
 * cell for the values down the left column. Exactly one input means a
 * one-variable table (its orientation follows which input is given); both
 * mean a two-variable grid:
 *
 * ```text
 * column          row             grid
 * · =formula      · v1 v2 v3      =formula r1 r2
 * v1 ?           =formula ? ? ?   c1  ?    ?
 * v2 ?                            c2  ?    ?
 * ```
 */
export function planDataTable(
  bounds: RangeBounds,
  rowInput: CellCoordinates | null,
  columnInput: CellCoordinates | null,
): DataTablePlan | DataTablePlanError {
  if (!rowInput && !columnInput) return 'noInputCell'
  if (
    rowInput &&
    columnInput &&
    rowInput.row === columnInput.row &&
    rowInput.column === columnInput.column
  ) {
    return 'sameInputCell'
  }
  for (const input of [rowInput, columnInput]) {
    if (input && inside(bounds, input)) return 'inputInsideTable'
  }
  const rows = bounds.endRow - bounds.startRow + 1
  const columns = bounds.endColumn - bounds.startColumn + 1

  let plan: DataTablePlan
  if (rowInput && columnInput) {
    // Two variables: the formula sits in the corner, the top row feeds the
    // row input, the left column feeds the column input.
    if (rows < 2 || columns < 2) return 'badShape'
    const rowValueCells = cellsAcrossRow(bounds)
    const columnValueCells = cellsDownColumn(bounds)
    plan = {
      kind: 'grid',
      formulaCell: { row: bounds.startRow, column: bounds.startColumn },
      rowValueCells,
      columnValueCells,
      resultStart: { row: bounds.startRow + 1, column: bounds.startColumn + 1 },
      resultRows: rows - 1,
      resultColumns: columns - 1,
    }
  } else if (columnInput) {
    // One variable, values down the left column; the formula is one row
    // above the first value and one column to its right.
    if (rows < 2 || columns !== 2) return 'badShape'
    const columnValueCells = cellsDownColumn(bounds)
    plan = {
      kind: 'column',
      formulaCell: { row: bounds.startRow, column: bounds.endColumn },
      rowValueCells: [],
      columnValueCells,
      resultStart: { row: bounds.startRow + 1, column: bounds.endColumn },
      resultRows: rows - 1,
      resultColumns: 1,
    }
  } else if (rowInput) {
    // One variable, values across the top row; the formula is one column
    // left of the first value and one row below it.
    if (columns < 2 || rows !== 2) return 'badShape'
    const rowValueCells = cellsAcrossRow(bounds)
    plan = {
      kind: 'row',
      formulaCell: { row: bounds.endRow, column: bounds.startColumn },
      rowValueCells,
      columnValueCells: [],
      resultStart: { row: bounds.endRow, column: bounds.startColumn + 1 },
      resultRows: 1,
      resultColumns: columns - 1,
    }
  } else {
    // Unreachable: the first check rejects the no-input case, but the
    // definite-assignment analysis cannot see that across the chain.
    return 'noInputCell'
  }
  if (plan.resultRows * plan.resultColumns > MAX_DATA_TABLE_CELLS) return 'tooManyCells'
  return plan
}

/// Runs `mutate` and resolves after the recalc it triggers (Goal Seek's
/// helper, kept local so the two what-if tools cannot drift).
function mutateAndAwaitRecalc(runtime: UniverRuntime, mutate: () => void): Promise<void> {
  const formula = runtime.univerAPI.getFormula()
  return new Promise((resolve) => {
    let settled = false
    let timer: ReturnType<typeof setTimeout> | null = null
    const finish = (): void => {
      if (settled) return
      settled = true
      disposable.dispose()
      if (timer) clearTimeout(timer)
      resolve()
    }
    const disposable = formula.calculationEnd(finish)
    timer = setTimeout(finish, 2_000)
    mutate()
    if (isManualCalculation(runtime)) calculateNow(runtime)
  })
}

export interface DataTableInput {
  /// The whole table region in A1 notation, e.g. "B4:D8" — corner, value
  /// headers and the (existing or empty) result cells.
  readonly tableRange: string
  /// Substitution target for the values across the table's top row.
  readonly rowInputCell: string | null
  /// Substitution target for the values down the table's left column.
  readonly columnInputCell: string | null
  /// Called once per evaluated cell, after the read (never during a write).
  readonly onProgress?: (done: number, total: number) => void
}

export interface DataTableOutcome {
  readonly kind: DataTableKind
  /// Number of result cells written.
  readonly cells: number
}

type ScalarValue = number | string | boolean | null

const errorMessage = (reason: DataTablePlanError): string => {
  if (reason === 'noInputCell') return t('dlgDataTableNeedInput')
  if (reason === 'sameInputCell') return t('dlgDataTableSameInput')
  if (reason === 'inputInsideTable') return t('dlgDataTableInputInside')
  if (reason === 'tooManyCells')
    return t('dlgDataTableTooManyCells', { limit: MAX_DATA_TABLE_CELLS })
  return t('dlgDataTableBadShape')
}

/**
 * Evaluates and writes a Data Table. Throws a user-facing Error on any
 * validation failure; the input cells keep (or are restored to) their
 * original contents in every case.
 */
export async function applyDataTable(
  runtime: UniverRuntime,
  input: DataTableInput,
): Promise<DataTableOutcome> {
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
  if (!worksheet) throw new Error(t('appNoActiveSheet'))

  let bounds: RangeBounds
  let rowInput: CellCoordinates | null = null
  let columnInput: CellCoordinates | null = null
  try {
    bounds = parseRange(input.tableRange.trim().toUpperCase())
    if (input.rowInputCell) rowInput = parseAddress(input.rowInputCell.trim().toUpperCase())
    if (input.columnInputCell)
      columnInput = parseAddress(input.columnInputCell.trim().toUpperCase())
  } catch {
    throw new Error(t('dlgDataTableBadRange'))
  }
  const plan = planDataTable(bounds, rowInput, columnInput)
  if (typeof plan === 'string') throw new Error(errorMessage(plan))

  const formulaRange = worksheet.getRange(plan.formulaCell.row, plan.formulaCell.column, 1, 1)
  if (!formulaRange.getFormula()) throw new Error(t('dlgDataTableNeedsFormula'))

  /// Excel's constraint too (Goal Seek's): the input cells must hold plain
  /// numbers (or nothing) — formulas or text would be destroyed by the run's
  /// writes, and the restore puts the original contents back at the end.
  type InputCell = { range: ReturnType<UniverWorksheet['getRange']> }
  const inputRanges: InputCell[] = []
  if (rowInput) inputRanges.push({ range: worksheet.getRange(rowInput.row, rowInput.column, 1, 1) })
  if (columnInput) {
    inputRanges.push({ range: worksheet.getRange(columnInput.row, columnInput.column, 1, 1) })
  }
  const originals: (number | null)[] = []
  for (const { range } of inputRanges) {
    const raw = range.getValue()
    if (range.getFormula() || (raw != null && typeof raw !== 'number')) {
      throw new Error(t('dlgDataTableInputNotNumber'))
    }
    originals.push(typeof raw === 'number' ? raw : null)
  }

  /// Substitution values must be numbers; anything else (text, blanks) would
  /// make most formulas error out for reasons the user cannot see in the table.
  const readNumbers = (cells: readonly CellCoordinates[]): number[] => {
    const values: number[] = []
    for (const cell of cells) {
      const raw = worksheet.getRange(cell.row, cell.column, 1, 1).getValue()
      if (typeof raw !== 'number') throw new Error(t('dlgDataTableValuesNotNumber'))
      values.push(raw)
    }
    return values
  }
  const rowValues = readNumbers(plan.rowValueCells)
  const columnValues = readNumbers(plan.columnValueCells)
  const total = plan.resultRows * plan.resultColumns

  const results: ScalarValue[][] = []
  let done = 0
  const report = (): void => {
    done += 1
    input.onProgress?.(done, total)
  }
  const undoBatch = beginUndoBatch(runtime)
  const restoreInputs = async (): Promise<void> => {
    for (let index = 0; index < inputRanges.length; index += 1) {
      const { range } = inputRanges[index]!
      const original = originals[index] ?? null
      await mutateAndAwaitRecalc(runtime, () =>
        original === null ? range.setValue({ v: null }) : range.setValue(original),
      )
    }
  }
  try {
    const readFormula = (): ScalarValue => {
      const value = formulaRange.getValue()
      return typeof value === 'number' || typeof value === 'string' || typeof value === 'boolean'
        ? value
        : null
    }
    if (plan.kind === 'grid') {
      const rowInputRange = inputRanges[0]!.range
      const columnInputRange = inputRanges[1]!.range
      for (const columnValue of columnValues) {
        const row: ScalarValue[] = []
        await mutateAndAwaitRecalc(runtime, () => columnInputRange.setValue(columnValue))
        for (const rowValue of rowValues) {
          await mutateAndAwaitRecalc(runtime, () => rowInputRange.setValue(rowValue))
          row.push(readFormula())
          report()
        }
        results.push(row)
      }
    } else {
      const axisRange = inputRanges[0]!.range
      const values = plan.kind === 'column' ? columnValues : rowValues
      const evaluated: ScalarValue[] = []
      for (const value of values) {
        await mutateAndAwaitRecalc(runtime, () => axisRange.setValue(value))
        evaluated.push(readFormula())
        report()
      }
      // The block shape differs: column tables stack results N×1, row
      // tables lay them out 1×M — setValues writes the matrix as given.
      results.push(...(plan.kind === 'column' ? evaluated.map((value) => [value]) : [evaluated]))
    }

    /// One journaled block write for the whole result area — errors read out
    /// of the formula (#DIV/0! mid-run is normal) travel as their text, and
    /// a blank formula cell writes as an explicit `{ v: null }` clear.
    const target = worksheet.getRange(
      plan.resultStart.row,
      plan.resultStart.column,
      plan.resultRows,
      plan.resultColumns,
    )
    await mutateAndAwaitRecalc(runtime, () =>
      target.setValues(results.map((row) => row.map((value) => ({ v: value })))),
    )
    await restoreInputs()
    return { kind: plan.kind, cells: total }
  } catch (error) {
    // A run cut short must not leave a substituted value in the grid; the
    // undo batch still holds every write done so far (one ⌘Z clears them).
    await restoreInputs()
    throw error
  } finally {
    undoBatch.settle()
  }
}
