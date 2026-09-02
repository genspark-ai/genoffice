/**
 * Excel-semantics Ctrl/Cmd+Arrow (and +Shift) data-edge navigation.
 *
 * Univer's jump navigation decides "does this cell have content" from the
 * computed display value only (`cellHasValue`: v non-empty or rich text p).
 * Excel stops at any cell with a value OR A FORMULA — so upstream skips
 * formula cells whose result is an empty string, and, worse, in cache mode
 * (values not materialized client-side) whole formula regions look blank and
 * Ctrl+Down sails past them to the sheet end (alpha: Merrick, v0.8.684).
 *
 * Fix: re-register the eight jump shortcuts (move + expand × 4 directions)
 * at a higher priority with the same preconditions, backed by a faithful
 * port of Univer's findNextGapRange whose emptiness test also counts
 * formulas (f / shared-formula si) and any present value including "".
 * Plain arrow keys and every other selection command stay upstream; when
 * the precondition fails (cell editor open, formula bar focused) the
 * dispatcher falls through to Univer's own items untouched.
 */
import {
  CommandType,
  Direction,
  ICommandService,
  IUniverInstanceService,
  RANGE_TYPE,
  Rectangle,
  getReverseDirection,
} from '@univerjs/core'
import type { ICellData, ICommand, IRange, Worksheet } from '@univerjs/core'
import { IShortcutService, KeyCode, MetaKeys } from '@univerjs/ui'
import type { IShortcutItem } from '@univerjs/ui'
import {
  SelectionMoveType,
  SetSelectionsOperation,
  SheetsSelectionsService,
  alignToMergedCellsBorders,
  getCellAtRowCol,
  getSheetCommandTarget,
} from '@univerjs/sheets'
import type { ISelectionWithStyle } from '@univerjs/sheets'
import { whenSheetEditorFocused } from '@univerjs/sheets-ui'
import type { UniverRuntime } from './univer-state'

const MOVE_JUMP_COMMAND_ID = 'genoffice.command.move-selection-excel-jump'
const EXPAND_JUMP_COMMAND_ID = 'genoffice.command.expand-selection-excel-jump'

interface IJumpParams {
  direction: Direction
}

type MatrixCell = ICellData & { rowSpan?: number; colSpan?: number }

/** Excel's rule: an entry is any value (including ""), rich text, or a formula */
function excelCellHasValue(cell: MatrixCell | null | undefined): boolean {
  if (cell == null) return false
  if (cell.v !== undefined && cell.v !== null) return true
  if (cell.p !== undefined && cell.p !== null) return true
  const f = (cell as { f?: unknown }).f
  const si = (cell as { si?: unknown }).si
  return (typeof f === 'string' && f.length > 0) || (typeof si === 'string' && si.length > 0)
}

function rangeHasValue(
  worksheet: Worksheet,
  row: number,
  col: number,
  rowEnd: number,
  colEnd: number,
): { hasValue: boolean; matrix: ReturnType<Worksheet['getMatrixWithMergedCells']> } {
  let hasValue = false
  const matrix = worksheet
    .getMatrixWithMergedCells(row, col, rowEnd, colEnd)
    .forValue((_, __, value) => {
      if (excelCellHasValue(value as MatrixCell)) {
        hasValue = true
        return false
      }
    })
  return { hasValue, matrix }
}

function getEdgeOfRange(startRange: IRange, direction: Direction, worksheet: Worksheet): IRange {
  let destRange: IRange
  switch (direction) {
    case Direction.UP:
      destRange = {
        startRow: startRange.startRow,
        startColumn: startRange.startColumn,
        endRow: startRange.startRow,
        endColumn: startRange.endColumn,
        rangeType: RANGE_TYPE.NORMAL,
      }
      break
    case Direction.DOWN:
      destRange = {
        startRow: startRange.endRow,
        startColumn: startRange.startColumn,
        endRow: startRange.endRow,
        endColumn: startRange.endColumn,
        rangeType: RANGE_TYPE.NORMAL,
      }
      break
    case Direction.LEFT:
      destRange = {
        startRow: startRange.startRow,
        startColumn: startRange.startColumn,
        endRow: startRange.endRow,
        endColumn: startRange.startColumn,
        rangeType: RANGE_TYPE.NORMAL,
      }
      break
    case Direction.RIGHT:
      destRange = {
        startRow: startRange.startRow,
        startColumn: startRange.endColumn,
        endRow: startRange.endRow,
        endColumn: startRange.endColumn,
        rangeType: RANGE_TYPE.NORMAL,
      }
      break
    default:
      throw new Error('Invalid direction')
  }
  return alignToMergedCellsBorders(destRange, worksheet, false)
}

/**
 * Faithful port of Univer's findNextGapRange with the Excel emptiness test.
 * Walks whole rows/columns of the (merge-aligned) edge range towards
 * `direction`, stopping at data-block boundaries the way Ctrl+Arrow does.
 */
function findNextGapRange(startRange: IRange, direction: Direction, worksheet: Worksheet): IRange {
  const destRange = { ...startRange }
  const { startRow, startColumn, endRow, endColumn } = getEdgeOfRange(
    startRange,
    direction,
    worksheet,
  )
  let currentPositionHasValue = rangeHasValue(
    worksheet,
    startRow,
    startColumn,
    endRow,
    endColumn,
  ).hasValue
  let firstMove = true
  let shouldContinue = true
  while (shouldContinue) {
    if (Direction.UP === direction) {
      let next = destRange.startRow - 1
      while (next > -1 && !worksheet.getRowVisible(next)) next -= 1
      if (next === -1) break
      const { hasValue: nextRangeHasValue, matrix } = rangeHasValue(
        worksheet,
        next,
        destRange.startColumn,
        next,
        destRange.endColumn,
      )
      if (currentPositionHasValue && !nextRangeHasValue && !firstMove) break
      if (matrix.getLength() !== 0) {
        let min = next
        matrix.forValue((row) => {
          min = Math.min(row, min)
        })
        destRange.startRow = min
      } else {
        destRange.startRow = next
      }
      destRange.endRow = destRange.startRow
      if (!currentPositionHasValue && nextRangeHasValue) break
      currentPositionHasValue = nextRangeHasValue
      firstMove = false
    } else if (Direction.DOWN === direction) {
      let next = destRange.endRow + 1
      while (next < worksheet.getRowCount() && !worksheet.getRowVisible(next)) next += 1
      if (next === worksheet.getRowCount()) break
      const { hasValue: nextRangeHasValue, matrix } = rangeHasValue(
        worksheet,
        next,
        destRange.startColumn,
        next,
        destRange.endColumn,
      )
      if (currentPositionHasValue && !nextRangeHasValue && !firstMove) break
      if (matrix.getLength() !== 0) {
        let max = next
        matrix.forValue((row, _, value) => {
          max = Math.max(row + ((value as MatrixCell).rowSpan || 1) - 1, max)
        })
        destRange.endRow = max
      } else {
        destRange.endRow = next
      }
      destRange.startRow = destRange.endRow
      if (!currentPositionHasValue && nextRangeHasValue) break
      currentPositionHasValue = nextRangeHasValue
      firstMove = false
    } else if (Direction.LEFT === direction) {
      let next = destRange.startColumn - 1
      while (next > -1 && !worksheet.getColVisible(next)) next -= 1
      if (next === -1) break
      const { hasValue: nextRangeHasValue, matrix } = rangeHasValue(
        worksheet,
        destRange.startRow,
        next,
        destRange.endRow,
        next,
      )
      if (currentPositionHasValue && !nextRangeHasValue && !firstMove) break
      if (matrix.getLength() !== 0) {
        let min = next
        matrix.forValue((_, col) => {
          min = Math.min(col, min)
        })
        destRange.startColumn = min
      } else {
        destRange.startColumn = next
      }
      destRange.endColumn = destRange.startColumn
      if (!currentPositionHasValue && nextRangeHasValue) break
      currentPositionHasValue = nextRangeHasValue
      firstMove = false
    } else if (Direction.RIGHT === direction) {
      let next = destRange.endColumn + 1
      while (next < worksheet.getColumnCount() && !worksheet.getColVisible(next)) next += 1
      if (next === worksheet.getColumnCount()) break
      const { hasValue: nextRangeHasValue, matrix } = rangeHasValue(
        worksheet,
        destRange.startRow,
        next,
        destRange.endRow,
        next,
      )
      if (currentPositionHasValue && !nextRangeHasValue && !firstMove) break
      if (matrix.getLength() !== 0) {
        let max = next
        matrix.forValue((_, col, value) => {
          max = Math.max(col + ((value as MatrixCell).colSpan || 1) - 1, max)
        })
        destRange.endColumn = max
      } else {
        destRange.endColumn = next
      }
      destRange.startColumn = destRange.endColumn
      if (!currentPositionHasValue && nextRangeHasValue) break
      currentPositionHasValue = nextRangeHasValue
      firstMove = false
    } else {
      shouldContinue = false
    }
  }
  return destRange
}

/** anchor the walk on the selection's primary (active) cell, as upstream does */
function getStartRange(
  range: IRange,
  primary: { actualRow: number; actualColumn: number } | null | undefined,
  direction: Direction,
): IRange {
  const ret = Rectangle.clone(range)
  if (primary == null) return ret
  switch (direction) {
    case Direction.UP:
    case Direction.DOWN:
      ret.startColumn = ret.endColumn = primary.actualColumn
      break
    case Direction.LEFT:
    case Direction.RIGHT:
      ret.startRow = ret.endRow = primary.actualRow
      break
  }
  return ret
}

function expandToNextGapRange(
  startRange: IRange,
  direction: Direction,
  worksheet: Worksheet,
): IRange {
  const next = findNextGapRange(startRange, direction, worksheet)
  return alignToMergedCellsBorders(Rectangle.union(next, startRange), worksheet, true)
}

function shrinkToNextGapRange(
  startRange: IRange,
  anchorRange: IRange,
  direction: Direction,
  worksheet: Worksheet,
): IRange {
  const nextGap = findNextGapRange(
    getEdgeOfRange(startRange, getReverseDirection(direction), worksheet),
    direction,
    worksheet,
  )
  if (direction === Direction.UP && nextGap.startRow <= startRange.startRow) {
    return alignToMergedCellsBorders(
      { ...anchorRange, startColumn: startRange.startColumn, endColumn: startRange.endColumn },
      worksheet,
      true,
    )
  }
  if (direction === Direction.DOWN && nextGap.endRow >= startRange.endRow) {
    return alignToMergedCellsBorders(
      { ...anchorRange, startColumn: startRange.startColumn, endColumn: startRange.endColumn },
      worksheet,
      true,
    )
  }
  if (direction === Direction.LEFT && nextGap.startColumn <= startRange.startColumn) {
    return alignToMergedCellsBorders(
      { ...anchorRange, startRow: startRange.startRow, endRow: startRange.endRow },
      worksheet,
      true,
    )
  }
  if (direction === Direction.RIGHT && nextGap.endColumn >= startRange.endColumn) {
    return alignToMergedCellsBorders(
      { ...anchorRange, startRow: startRange.startRow, endRow: startRange.endRow },
      worksheet,
      true,
    )
  }
  return Rectangle.union(Rectangle.clone(anchorRange), nextGap)
}

function checkIfShrink(
  selection: ISelectionWithStyle,
  direction: Direction,
  worksheet: Worksheet,
): boolean {
  const { primary, range } = selection
  const startRange = Rectangle.clone(range)
  switch (direction) {
    case Direction.UP:
    case Direction.DOWN:
      startRange.startRow = primary?.startRow ?? range.startRow
      startRange.endRow = primary?.endRow ?? range.startRow
      break
    case Direction.LEFT:
    case Direction.RIGHT:
      startRange.startColumn = primary?.startColumn ?? range.startColumn
      startRange.endColumn = primary?.endColumn ?? range.startColumn
      break
  }
  const anchorRange = getEdgeOfRange(startRange, direction, worksheet)
  switch (direction) {
    case Direction.DOWN:
      return range.startRow < anchorRange.startRow
    case Direction.UP:
      return range.endRow > anchorRange.endRow
    case Direction.LEFT:
      return anchorRange.endColumn < range.endColumn
    case Direction.RIGHT:
      return anchorRange.startColumn > range.startColumn
    default:
      return false
  }
}

export function registerExcelJumpNav(runtime: UniverRuntime): void {
  const injector = runtime.univer.__getInjector()
  const commandService = injector.get(ICommandService)
  const shortcutService = injector.get(IShortcutService)

  const moveCommand: ICommand<IJumpParams> = {
    id: MOVE_JUMP_COMMAND_ID,
    type: CommandType.COMMAND,
    handler: (accessor, params) => {
      if (!params) return false
      const target = getSheetCommandTarget(accessor.get(IUniverInstanceService))
      if (!target) return false
      const { workbook, worksheet } = target
      const selection = accessor.get(SheetsSelectionsService).getCurrentLastSelection()
      if (!selection) return false
      const { direction } = params
      const { range, primary } = selection
      const startRange = getStartRange(range, primary, direction)
      const next = findNextGapRange(startRange, direction, worksheet)
      const destRange = getCellAtRowCol(next.startRow, next.startColumn, worksheet)
      if (Rectangle.equals(destRange, startRange)) return false
      return accessor.get(ICommandService).executeCommand(SetSelectionsOperation.id, {
        unitId: workbook.getUnitId(),
        subUnitId: worksheet.getSheetId(),
        selections: [
          {
            range: Rectangle.clone(destRange),
            primary: {
              startRow: destRange.startRow,
              startColumn: destRange.startColumn,
              endRow: destRange.endRow,
              endColumn: destRange.endColumn,
              actualRow: next.startRow,
              actualColumn: next.startColumn,
              isMerged: destRange.isMerged,
              isMergedMainCell:
                destRange.startRow === next.startRow && destRange.startColumn === next.startColumn,
            },
          },
        ],
        type: SelectionMoveType.MOVE_END,
        reveal: true,
      })
    },
  }

  const expandCommand: ICommand<IJumpParams> = {
    id: EXPAND_JUMP_COMMAND_ID,
    type: CommandType.COMMAND,
    handler: (accessor, params) => {
      if (!params) return false
      const target = getSheetCommandTarget(accessor.get(IUniverInstanceService))
      if (!target) return false
      const { worksheet, unitId, subUnitId } = target
      const selection = accessor.get(SheetsSelectionsService).getCurrentLastSelection()
      if (!selection) return false
      const { direction } = params
      const { range: startRange, primary } = selection
      const destRange = !checkIfShrink(selection, direction, worksheet)
        ? expandToNextGapRange(startRange, direction, worksheet)
        : shrinkToNextGapRange(
            startRange,
            { ...Rectangle.clone(primary as IRange), rangeType: RANGE_TYPE.NORMAL },
            direction,
            worksheet,
          )
      if (selection.range.rangeType !== undefined) destRange.rangeType = selection.range.rangeType
      if (Rectangle.equals(destRange, startRange)) return false
      return accessor.get(ICommandService).syncExecuteCommand(SetSelectionsOperation.id, {
        unitId,
        subUnitId,
        type: SelectionMoveType.MOVE_END,
        selections: [{ range: destRange, primary }],
        reveal: true,
      })
    },
  }

  // registrations live for the app lifetime (one workbook per window)
  commandService.registerCommand(moveCommand)
  commandService.registerCommand(expandCommand)

  const directions: Array<[Direction, KeyCode]> = [
    [Direction.DOWN, KeyCode.ARROW_DOWN],
    [Direction.UP, KeyCode.ARROW_UP],
    [Direction.LEFT, KeyCode.ARROW_LEFT],
    [Direction.RIGHT, KeyCode.ARROW_RIGHT],
  ]
  for (const [direction, keyCode] of directions) {
    const move: IShortcutItem<IJumpParams> = {
      id: MOVE_JUMP_COMMAND_ID,
      // outranks Univer's own items on the same binding; when the
      // precondition fails the dispatcher falls through to upstream
      priority: 100,
      binding: keyCode | MetaKeys.CTRL_COMMAND,
      preconditions: whenSheetEditorFocused,
      staticParameters: { direction },
    }
    const expand: IShortcutItem<IJumpParams> = {
      id: EXPAND_JUMP_COMMAND_ID,
      priority: 100,
      binding: keyCode | MetaKeys.CTRL_COMMAND | MetaKeys.SHIFT,
      preconditions: whenSheetEditorFocused,
      staticParameters: { direction },
    }
    shortcutService.registerShortcut(move)
    shortcutService.registerShortcut(expand)
  }
}

/** exported for tests */
export const _internals = {
  excelCellHasValue,
  findNextGapRange,
  MOVE_JUMP_COMMAND_ID,
  EXPAND_JUMP_COMMAND_ID,
}
