/**
 * Sparse scanning for Univer's sheets find provider.
 *
 * The built-in SheetFindModel walks every coordinate of rowCount × columnCount
 * through the getCell interceptor chain (conditional formats, number formats,
 * formulas — ~15µs a cell). A grid that is wide or tall but sparse — a
 * <dimension> declared out to XFD, a stray cell far below the data, a streamed
 * workbook's full row count — turns each keystroke in the Find dialog into
 * seconds of frozen renderer (half a million empty cells for a 32-row sheet).
 *
 * This narrows every range the model scans to the rows (or columns, for
 * column-direction searches) that hold raw cells or array-formula spills,
 * spanning only their populated extent, and hands those slices to Univer's own
 * iterator. Hit semantics and match order are unchanged: the dense walk skips
 * empty cells anyway, and the slices are visited in the same axis order.
 */
import type { IRange, Worksheet } from '@univerjs/core'
import { FindDirection, type IFindQuery } from '@univerjs/find-replace'

type FindRangeResult = { results: unknown[] }
type FindInRange = (
  this: unknown,
  worksheet: Worksheet,
  query: IFindQuery,
  range: IRange,
  unitId: string,
  dedupeFn?: (row: number, col: number) => boolean,
) => FindRangeResult

interface FindModelPrototype {
  _findInRange?: FindInRange
  [SPARSE_PATCHED]?: true
}

interface Injector {
  createInstance(ctor: unknown, ...args: unknown[]): unknown
}

/** Where array formulas spill values that never enter the raw cell matrix. */
export type SpillLookup = (unitId: string, sheetId: string) => IRange[]

const SPARSE_PATCHED = Symbol('sparse-find-patched')

/** Populated extent of one row (or column) on the other axis. */
export interface AxisRun {
  index: number
  start: number
  end: number
}

interface RawMatrix {
  [row: string]: { [col: string]: unknown } | undefined
}

function extend(runs: Map<number, AxisRun>, index: number, position: number): void {
  const run = runs.get(index)
  if (!run) runs.set(index, { index, start: position, end: position })
  else {
    if (position < run.start) run.start = position
    if (position > run.end) run.end = position
  }
}

/**
 * Populated runs of a worksheet: per row (default) or per column, sorted by
 * axis index. Raw cells and spill ranges both count; everything else is empty
 * for a search and never needs a getCell call.
 */
export function collectRuns(matrix: RawMatrix, spills: IRange[], byColumn: boolean): AxisRun[] {
  const runs = new Map<number, AxisRun>()
  for (const rowKey of Object.keys(matrix)) {
    const rowCells = matrix[rowKey]
    if (!rowCells) continue
    const row = Number(rowKey)
    for (const colKey of Object.keys(rowCells)) {
      if (rowCells[colKey] === undefined || rowCells[colKey] === null) continue
      const col = Number(colKey)
      if (byColumn) extend(runs, col, row)
      else extend(runs, row, col)
    }
  }
  for (const spill of spills) {
    if (byColumn) {
      for (let col = spill.startColumn; col <= spill.endColumn; col += 1) {
        extend(runs, col, spill.startRow)
        extend(runs, col, spill.endRow)
      }
    } else {
      for (let row = spill.startRow; row <= spill.endRow; row += 1) {
        extend(runs, row, spill.startColumn)
        extend(runs, row, spill.endColumn)
      }
    }
  }
  return [...runs.values()].sort((a, b) => a.index - b.index)
}

/** Intersects the runs with `range`, yielding one single-row/column slice each. */
export function sliceRange(range: IRange, runs: AxisRun[], byColumn: boolean): IRange[] {
  const [axisStart, axisEnd, otherStart, otherEnd] = byColumn
    ? [range.startColumn, range.endColumn, range.startRow, range.endRow]
    : [range.startRow, range.endRow, range.startColumn, range.endColumn]
  const slices: IRange[] = []
  for (const run of runs) {
    if (run.index < axisStart || run.index > axisEnd) continue
    const start = Math.max(run.start, otherStart)
    const end = Math.min(run.end, otherEnd)
    if (start > end) continue
    slices.push(
      byColumn
        ? { startRow: start, endRow: end, startColumn: run.index, endColumn: run.index }
        : { startRow: run.index, endRow: run.index, startColumn: start, endColumn: end },
    )
  }
  return slices
}

function rawMatrix(worksheet: Worksheet): RawMatrix {
  return worksheet.getCellMatrix().getMatrix() as unknown as RawMatrix
}

/** Replaces `_findInRange` on the model prototype with the sliced walk. */
export function patchFindModel(proto: FindModelPrototype, spills: SpillLookup): boolean {
  if (proto[SPARSE_PATCHED]) return true
  const original = proto._findInRange
  if (typeof original !== 'function') return false
  proto._findInRange = function sparseFindInRange(worksheet, query, range, unitId, dedupeFn) {
    const byColumn = query.findDirection === FindDirection.COLUMN
    let runs: AxisRun[]
    try {
      runs = collectRuns(rawMatrix(worksheet), spills(unitId, worksheet.getSheetId()), byColumn)
    } catch {
      return original.call(this, worksheet, query, range, unitId, dedupeFn)
    }
    const results: unknown[] = []
    for (const slice of sliceRange(range, runs, byColumn)) {
      for (const hit of original.call(this, worksheet, query, slice, unitId, dedupeFn).results) {
        results.push(hit)
      }
    }
    return { results }
  }
  proto[SPARSE_PATCHED] = true
  return true
}

/**
 * The model class is private to the provider and only ever reaches us through
 * `provider._injector.createInstance(SheetFindModel, …)` inside find(). Wrap
 * that injector so the prototype is patched before the first model is built —
 * the first search is the one the user is waiting on. Unknown provider shapes
 * are left alone (the dense walk still works, just slowly).
 */
export function installSparseFind(provider: unknown, spills: SpillLookup): boolean {
  const host = provider as { _injector?: Injector; [SPARSE_PATCHED]?: true }
  if (host[SPARSE_PATCHED]) return true
  const injector = host._injector
  if (!injector || typeof injector.createInstance !== 'function') return false
  host._injector = new Proxy(injector, {
    get(target, property) {
      if (property === 'createInstance') {
        return (ctor: unknown, ...args: unknown[]) => {
          const proto = (ctor as { prototype?: FindModelPrototype } | null)?.prototype
          if (proto) patchFindModel(proto, spills)
          return target.createInstance(ctor, ...args)
        }
      }
      const value = Reflect.get(target, property, target)
      return typeof value === 'function' ? value.bind(target) : value
    },
  })
  host[SPARSE_PATCHED] = true
  return true
}
