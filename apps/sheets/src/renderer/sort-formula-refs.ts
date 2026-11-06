import {
  Range,
  type ICellData,
  type IRange,
  type Nullable,
  type ObjectMatrix,
} from '@univerjs/core'
import { LexerTreeBuilder } from '@univerjs/engine-formula'
import {
  ReorderRangeCommand,
  SetRangeValuesMutation,
  SheetInterceptorService,
  type IReorderRangeCommandParams,
} from '@univerjs/sheets'

import type { UniverRuntime } from './univer-state'

type CellPatch = Record<number, Record<number, ICellData>>

function cutGroupPatches(
  matrix: ObjectMatrix<Nullable<ICellData>>,
  range: IRange,
  moveRefOffset: (formula: string, dx: number, dy: number) => string,
): { redo: CellPatch; undo: CellPatch } | null {
  const groups = new Set<string>()
  Range.foreach(range, (row, column) => {
    const cell = matrix.getValue(row, column)
    if (cell?.f && cell.si) groups.add(cell.si)
  })
  if (groups.size === 0) return null
  const inRange = (row: number, column: number) =>
    row >= range.startRow &&
    row <= range.endRow &&
    column >= range.startColumn &&
    column <= range.endColumn
  const masters = new Map<string, { row: number; column: number; f: string }>()
  const outside: { row: number; column: number; f: Nullable<string>; si: string }[] = []
  matrix.forValue((row, column, cell) => {
    if (!cell?.si || !groups.has(cell.si)) return
    if (cell.f && !masters.has(cell.si)) masters.set(cell.si, { row, column, f: cell.f })
    if (!inRange(row, column)) outside.push({ row, column, f: cell.f, si: cell.si })
  })
  const redo: CellPatch = {}
  const undo: CellPatch = {}
  let changed = false
  for (const { row, column, f, si } of outside) {
    const master = masters.get(si)
    if (!master || !inRange(master.row, master.column)) continue
    ;(redo[row] ??= {})[column] = {
      f: f || moveRefOffset(master.f, column - master.column, row - master.row),
      si: null,
    }
    ;(undo[row] ??= {})[column] = { f: f ?? null, si }
    changed = true
  }
  return changed ? { redo, undo } : null
}

export function installSortFormulaRefs(runtime: UniverRuntime): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const lexer = injector.get(LexerTreeBuilder)
  return injector.get(SheetInterceptorService).interceptCommand({
    getMutations: (command) => {
      if (command.id !== ReorderRangeCommand.id) return { redos: [], undos: [] }
      const { unitId, subUnitId, range } = command.params as IReorderRangeCommandParams
      const matrix = runtime.univerAPI
        .getWorkbook(unitId)
        ?.getSheetBySheetId(subUnitId)
        ?.getSheet()
        .getCellMatrix()
      if (!matrix) return { redos: [], undos: [] }
      const patches = cutGroupPatches(matrix, range, (formula, dx, dy) =>
        lexer.moveFormulaRefOffset(formula, dx, dy),
      )
      if (!patches) return { redos: [], undos: [] }
      return {
        redos: [
          { id: SetRangeValuesMutation.id, params: { unitId, subUnitId, cellValue: patches.redo } },
        ],
        undos: [
          { id: SetRangeValuesMutation.id, params: { unitId, subUnitId, cellValue: patches.undo } },
        ],
      }
    },
  })
}
