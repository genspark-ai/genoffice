/**
 * Excel's Insert Cells / Delete Cells (Ctrl+Shift++ / Ctrl+-): a whole-row
 * or whole-column selection inserts/removes those lines straight away; any
 * other selection asks how to shift the neighbours.
 */
import { IUniverInstanceService, RANGE_TYPE } from '@univerjs/core'
import type { IRange } from '@univerjs/core'
import { SheetsSelectionsService, getSheetCommandTarget } from '@univerjs/sheets'
import { isModalOpen } from './global-shortcuts'
import type { UniverRuntime } from './univer-state'

export type CellsMode = 'insert' | 'delete'
export type CellsChoice = 'shift-horizontal' | 'shift-vertical' | 'entire-row' | 'entire-column'

export const CELLS_CHOICES: readonly CellsChoice[] = [
  'shift-horizontal',
  'shift-vertical',
  'entire-row',
  'entire-column',
]

export const DEFAULT_CELLS_CHOICE: CellsChoice = 'shift-vertical'

export const INSERT_ROWS_COMMAND = 'sheet.command.insert-row-before'
export const INSERT_COLS_COMMAND = 'sheet.command.insert-col-before'
export const REMOVE_ROWS_COMMAND = 'sheet.command.remove-row'
export const REMOVE_COLS_COMMAND = 'sheet.command.remove-col'
export const INSERT_RANGE_RIGHT_COMMAND = 'sheet.command.insert-range-move-right'
export const INSERT_RANGE_DOWN_COMMAND = 'sheet.command.insert-range-move-down'
export const DELETE_RANGE_LEFT_COMMAND = 'sheet.command.delete-range-move-left'
export const DELETE_RANGE_UP_COMMAND = 'sheet.command.delete-range-move-up'

export function cellsCommandId(mode: CellsMode, choice: CellsChoice): string {
  switch (choice) {
    case 'entire-row':
      return mode === 'insert' ? INSERT_ROWS_COMMAND : REMOVE_ROWS_COMMAND
    case 'entire-column':
      return mode === 'insert' ? INSERT_COLS_COMMAND : REMOVE_COLS_COMMAND
    case 'shift-horizontal':
      return mode === 'insert' ? INSERT_RANGE_RIGHT_COMMAND : DELETE_RANGE_LEFT_COMMAND
    case 'shift-vertical':
      return mode === 'insert' ? INSERT_RANGE_DOWN_COMMAND : DELETE_RANGE_UP_COMMAND
  }
}

export interface SelectionShape {
  readonly range: IRange
  readonly rowCount: number
  readonly columnCount: number
}

export function isWholeRowSelection({ range, columnCount }: SelectionShape): boolean {
  return (
    range.rangeType === RANGE_TYPE.ROW ||
    range.rangeType === RANGE_TYPE.ALL ||
    (range.startColumn <= 0 && range.endColumn >= columnCount - 1)
  )
}

export function isWholeColumnSelection({ range, rowCount }: SelectionShape): boolean {
  return (
    range.rangeType === RANGE_TYPE.COLUMN ||
    range.rangeType === RANGE_TYPE.ALL ||
    (range.startRow <= 0 && range.endRow >= rowCount - 1)
  )
}

export type CellsDecision = { kind: 'direct'; choice: CellsChoice } | { kind: 'dialog' }

/** Whole rows/columns act at once like Excel; anything else needs the dialog. */
export function decideCellsAction(shape: SelectionShape): CellsDecision {
  if (isWholeRowSelection(shape)) return { kind: 'direct', choice: 'entire-row' }
  if (isWholeColumnSelection(shape)) return { kind: 'direct', choice: 'entire-column' }
  return { kind: 'dialog' }
}

/**
 * Univer's insert-before commands count from `value`, not the selection;
 * the remove and range-shift commands read the selection themselves.
 */
export function cellsCommandParams(
  mode: CellsMode,
  choice: CellsChoice,
  range: IRange,
): { value: number } | undefined {
  if (mode !== 'insert') return undefined
  if (choice === 'entire-row') return { value: range.endRow - range.startRow + 1 }
  if (choice === 'entire-column') return { value: range.endColumn - range.startColumn + 1 }
  return undefined
}

export function currentSelectionShape(runtime: UniverRuntime): SelectionShape | null {
  const injector = runtime.univer.__getInjector()
  const target = getSheetCommandTarget(injector.get(IUniverInstanceService))
  const range = injector.get(SheetsSelectionsService).getCurrentLastSelection()?.range
  if (!target || !range) return null
  return {
    range,
    rowCount: target.worksheet.getRowCount(),
    columnCount: target.worksheet.getColumnCount(),
  }
}

/// Dispatches the chosen action; the app's BeforeCommandExecute gates
/// (pivot sheets, streamed workbooks, formula spans) may still cancel it.
export function runCellsChoice(runtime: UniverRuntime, mode: CellsMode, choice: CellsChoice): void {
  const shape = currentSelectionShape(runtime)
  if (!shape) return
  void runtime.univerAPI.executeCommand(
    cellsCommandId(mode, choice),
    cellsCommandParams(mode, choice, shape.range),
  )
}

/** Shortcut / ribbon entry: acts directly or asks the host to open the dialog. */
export function requestCellsAction(
  runtime: UniverRuntime,
  mode: CellsMode,
  openDialog: (mode: CellsMode) => void,
): boolean {
  if (isModalOpen()) return false
  const shape = currentSelectionShape(runtime)
  if (!shape) return false
  const decision = decideCellsAction(shape)
  if (decision.kind === 'dialog') openDialog(mode)
  else runCellsChoice(runtime, mode, decision.choice)
  return true
}
