/**
 * Selection edge-wrap guard. Univer's MoveSelectionCommand calls
 * findNextRange with its default isGoBack=true, so an arrow key at the
 * sheet edge wraps the selection to the opposite side and shifts one
 * row/column (Z1000 + Down lands on A1's path). Enter/Tab have an
 * upstream no-wrap guard; arrow keys don't — cancel the command when no
 * visible cell exists in the pressed direction. Ctrl+arrow (jumpOver)
 * never wraps upstream, so the same cancel is a no-op for it.
 */
import { Direction } from '@univerjs/core'

import type { UniverRuntime } from './univer-state'

export const MOVE_SELECTION_COMMAND = 'sheet.command.move-selection'

export interface WorksheetGrid {
  getRowCount(): number
  getColumnCount(): number
  getRowVisible(row: number): boolean
  getColVisible(col: number): boolean
}

export interface SelectionRange {
  startRow: number
  startColumn: number
  endRow: number
  endColumn: number
}

export function moveWouldWrap(
  direction: Direction,
  range: SelectionRange,
  sheet: WorksheetGrid,
): boolean {
  switch (direction) {
    case Direction.UP:
      for (let row = range.startRow - 1; row >= 0; row -= 1) {
        if (sheet.getRowVisible(row)) return false
      }
      return true
    case Direction.DOWN:
      for (let row = range.endRow + 1; row < sheet.getRowCount(); row += 1) {
        if (sheet.getRowVisible(row)) return false
      }
      return true
    case Direction.LEFT:
      for (let col = range.startColumn - 1; col >= 0; col -= 1) {
        if (sheet.getColVisible(col)) return false
      }
      return true
    case Direction.RIGHT:
      for (let col = range.endColumn + 1; col < sheet.getColumnCount(); col += 1) {
        if (sheet.getColVisible(col)) return false
      }
      return true
    default:
      return false
  }
}

export function installSelectionWrapGuard(runtime: UniverRuntime): { dispose(): void } {
  return runtime.univerAPI.addEvent(runtime.univerAPI.Event.BeforeCommandExecute, (event) => {
    if (event.id !== MOVE_SELECTION_COMMAND) return
    const { direction } = (event.params ?? {}) as { direction?: Direction }
    if (direction === undefined) return
    const workbook = runtime.univerAPI.getActiveWorkbook()
    const range = workbook?.getActiveRange()?.getRange()
    const sheet = workbook?.getActiveSheet()?.getSheet()
    if (!range || !sheet) return
    if (moveWouldWrap(direction, range, sheet)) event.cancel = true
  })
}
