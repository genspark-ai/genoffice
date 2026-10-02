import { IRenderManagerService } from '@univerjs/engine-render'
import { SheetScrollManagerService, SheetSkeletonManagerService } from '@univerjs/sheets-ui'

import type { UniverRuntime } from './univer-state'

/// Grows a worksheet's grid as the viewport approaches its edge.
///
/// The grid a sheet is created with is sized to its data (see
/// MINIMUM_SHEET_COLUMN_COUNT / MINIMUM_SHEET_ROW_COUNT in univer-sync), so a
/// blank workbook stops being scrollable a few columns past the last cell. This
/// grows it while the user works instead, the way Excel does: scrolling right
/// keeps revealing columns, because reaching the edge extends the grid before
/// the next scroll would have stopped at it.
///
/// The growth is a view concern only. It writes no journal entry and never
/// reaches the snapshot, so the grid a file declares on save is still the used
/// range — the same split the engine already makes for a <dimension> that runs
/// to the last column (workbook.rs declared_extent).

/// Univer's last valid column and row index (0-based). The .xlsx grid caps at
/// 16384 x 1048576, so growth stops here rather than past it.
const LAST_COLUMN_INDEX = 16_383
const LAST_ROW_INDEX = 1_048_575

/// How close the viewport's first column/row gets to the edge before the grid
/// extends. The extension is only visible once the user reaches it, so a
/// viewport parked at the edge would otherwise sit against a dead border.
const COLUMN_LOOKAHEAD = 20
const ROW_LOOKAHEAD = 200

/// How much one extension adds. A column is a fraction of a screen wide and a
/// row is a couple of lines tall, so the axes need different steps to cover the
/// same distance.
const COLUMN_STEP = 26
const ROW_STEP = 200

export interface GridEdge {
  /// first visible column and row, 0-based (Univer's own indices)
  readonly startColumn: number
  readonly startRow: number
  /// the grid's present size, in columns and rows
  readonly columnCount: number
  readonly rowCount: number
}

export interface GridGrowth {
  readonly columnCount: number
  readonly rowCount: number
}

/**
 * The grid size the viewport at `edge` needs, or null when it already has one.
 *
 * Growth is self-limiting: the caller's new size puts the viewport back outside
 * the lookahead, so a burst of scroll events extends the grid once rather than
 * once per event.
 */
export function nextGridGrowth(edge: GridEdge): GridGrowth | null {
  let columnCount = edge.columnCount
  if (columnCount < LAST_COLUMN_INDEX + 1 && edge.startColumn + COLUMN_LOOKAHEAD >= columnCount) {
    columnCount = Math.min(columnCount + COLUMN_STEP, LAST_COLUMN_INDEX + 1)
  }
  let rowCount = edge.rowCount
  if (rowCount < LAST_ROW_INDEX + 1 && edge.startRow + ROW_LOOKAHEAD >= rowCount) {
    rowCount = Math.min(rowCount + ROW_STEP, LAST_ROW_INDEX + 1)
  }
  if (columnCount === edge.columnCount && rowCount === edge.rowCount) return null
  return { columnCount, rowCount }
}

interface ScrollState {
  readonly sheetViewStartColumn?: number
  readonly sheetViewStartRow?: number
}

interface FWorksheetLike {
  getMaxColumns(): number
  getMaxRows(): number
  setColumnCount(count: number): unknown
  setRowCount(count: number): unknown
}

interface SkeletonManagerLike {
  reCalculate(param?: unknown): void
}

/**
 * Grows the active sheet's grid as its viewport nears the edge.
 *
 * Returns a disposer. Missing render services (a runtime that never installed
 * sheets-ui) are not an error: the grid then stays at the size the data
 * implies, which is what it did before.
 */
export function installGridGrowth(runtime: UniverRuntime): () => void {
  let subscription: { unsubscribe(): void } | undefined
  let growing = false
  try {
    const injector = (
      runtime.univer as unknown as {
        __getInjector(): {
          get<T>(token: unknown): T
          get<T>(token: unknown, name: string): T
        }
      }
    ).__getInjector()
    const service = injector.get<{
      validViewportScrollInfo$?: {
        subscribe(next: (state: ScrollState | null) => void): { unsubscribe(): void }
      }
    }>(SheetScrollManagerService)
    const observable = service?.validViewportScrollInfo$
    if (!observable) return () => {}

    subscription = observable.subscribe((state) => {
      // setColumnCount re-lays the skeleton, which can emit another scroll
      // state synchronously. The new size already sits outside the lookahead so
      // the nested call returns null, but the flag keeps that from depending on
      // emission order.
      if (growing || !state) return
      growing = true
      try {
        const worksheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet() as
          FWorksheetLike | undefined
        if (!worksheet) return
        const growth = nextGridGrowth({
          startColumn: state.sheetViewStartColumn ?? 0,
          startRow: state.sheetViewStartRow ?? 0,
          columnCount: worksheet.getMaxColumns(),
          rowCount: worksheet.getMaxRows(),
        })
        if (!growth) return
        if (growth.columnCount !== worksheet.getMaxColumns()) {
          worksheet.setColumnCount(growth.columnCount)
        }
        if (growth.rowCount !== worksheet.getMaxRows()) {
          worksheet.setRowCount(growth.rowCount)
        }
        // The skeleton caches the row/column sizes it was built with, so a new
        // size is invisible until it is rebuilt.
        const render = (
          runtime.univer as unknown as {
            __getInjector(): { get<T>(token: unknown): T }
          }
        )
          .__getInjector()
          .get<{ getRenderById(id: string): { with<T>(token: unknown): T } | null } | null>(
            IRenderManagerService,
          )
        const unitId = runtime.univerAPI.getActiveWorkbook()?.getId()
        const skeleton = unitId ? render?.getRenderById(unitId) : null
        skeleton?.with<SkeletonManagerLike>(SheetSkeletonManagerService)?.reCalculate()
      } catch {
        // a sheet that is mid-rebuild has no skeleton to grow; the next scroll
        // state re-runs this with the rebuilt grid
      } finally {
        growing = false
      }
    })
  } catch {
    return () => {} // sheets-ui not installed in this runtime
  }
  return () => {
    subscription?.unsubscribe()
  }
}
