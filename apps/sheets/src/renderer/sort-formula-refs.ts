/**
 * A formula sorted to another row keeps pointing at its own row: Excel and
 * LibreOffice shift its relative references by the distance it moved, as a
 * copy does. Univer's reorder-range mutation copies the raw cells instead,
 * so D2's `=B2*C2` sorted to row 21 still reads B2:C2, and a shared-formula
 * (`si`) follower is re-derived from wherever its master landed.
 *
 * After each reorder-range, the moved formulas are rewritten with shifted
 * references, and every cell of a shared-formula group the sort touched
 * becomes the plain formula it showed before (shifted if it moved). The
 * undo restores the original cells.
 */
import { type ICellData, type IRange, type Nullable } from '@univerjs/core'
import { LexerTreeBuilder } from '@univerjs/engine-formula'
import {
  ReorderRangeCommand,
  SetRangeValuesMutation,
  SheetInterceptorService,
  type IReorderRangeCommandParams,
} from '@univerjs/sheets'

import type { UniverRuntime } from './univer-state'

interface FormulaCell {
  f?: Nullable<string>
  si?: Nullable<string>
}

interface ReorderSource {
  /** raw cell content at (row, column) before the reorder */
  getCell(row: number, column: number): Nullable<FormulaCell>
  /** every non-empty cell of the sheet, for shared-formula groups */
  forEachCell(visit: (row: number, column: number, cell: FormulaCell) => void): void
}

type CellPatch = Record<number, Record<number, ICellData>>

/**
 * The cell patches a reorder needs: `redo` applies after the reorder
 * mutation, `undo` after its inverse. `order` maps a target row to the row
 * whose content lands there.
 */
function sortFormulaPatches(
  source: ReorderSource,
  range: IRange,
  order: Readonly<Record<number, number>>,
  moveRefOffset: (formula: string, dx: number, dy: number) => string,
): { redo: CellPatch; undo: CellPatch } | null {
  const targetOf = new Map<number, number>()
  for (const [target, from] of Object.entries(order)) targetOf.set(from, Number(target))
  const inRange = (row: number, column: number) =>
    row >= range.startRow &&
    row <= range.endRow &&
    column >= range.startColumn &&
    column <= range.endColumn
  const landing = (row: number, column: number) =>
    inRange(row, column) ? (targetOf.get(row) ?? row) : row

  const masters = new Map<string, { row: number; column: number; f: string }>()
  const groupsTouched = new Set<string>()
  source.forEachCell((row, column, cell) => {
    if (!cell.si) return
    if (cell.f && !masters.has(cell.si)) masters.set(cell.si, { row, column, f: cell.f })
    if (landing(row, column) !== row) groupsTouched.add(cell.si)
  })
  const shownFormula = (row: number, column: number, cell: FormulaCell): string | null => {
    if (cell.f) return cell.f
    const master = cell.si ? masters.get(cell.si) : undefined
    if (!master) return null
    return moveRefOffset(master.f, column - master.column, row - master.row)
  }

  const redo: CellPatch = {}
  const undo: CellPatch = {}
  let changed = false
  const patch = (row: number, column: number, cell: FormulaCell) => {
    const formula = shownFormula(row, column, cell)
    if (formula === null) return
    const target = landing(row, column)
    const moved = target === row ? formula : moveRefOffset(formula, 0, target - row)
    if (moved === cell.f && !cell.si) return
    ;(redo[target] ??= {})[column] = { f: moved, si: null }
    ;(undo[row] ??= {})[column] = { f: cell.f ?? null, si: cell.si ?? null }
    changed = true
  }
  for (const [from, target] of targetOf) {
    if (from === target) continue
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const cell = source.getCell(from, column)
      if (cell && (cell.f || cell.si) && !(cell.si && groupsTouched.has(cell.si))) {
        patch(from, column, cell)
      }
    }
  }
  if (groupsTouched.size > 0) {
    source.forEachCell((row, column, cell) => {
      if (cell.si && groupsTouched.has(cell.si)) patch(row, column, cell)
    })
  }
  return changed ? { redo, undo } : null
}

export function installSortFormulaRefs(runtime: UniverRuntime): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const lexer = injector.get(LexerTreeBuilder)
  return injector.get(SheetInterceptorService).interceptCommand({
    getMutations: (command) => {
      if (command.id !== ReorderRangeCommand.id) return { redos: [], undos: [] }
      const { unitId, subUnitId, range, order } = command.params as IReorderRangeCommandParams
      const worksheet = runtime.univerAPI
        .getWorkbook(unitId)
        ?.getSheetBySheetId(subUnitId)
        ?.getSheet()
      if (!worksheet) return { redos: [], undos: [] }
      const matrix = worksheet.getCellMatrix()
      const patches = sortFormulaPatches(
        {
          getCell: (row, column) => matrix.getValue(row, column),
          forEachCell: (visit) =>
            matrix.forValue((row, column, cell) => {
              if (cell) visit(row, column, cell)
            }),
        },
        range,
        order,
        (formula, dx, dy) => lexer.moveFormulaRefOffset(formula, dx, dy),
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
