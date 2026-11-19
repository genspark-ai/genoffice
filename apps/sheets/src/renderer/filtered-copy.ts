/**
 * Copy semantics for filtered sheets: copying a range on a sheet
 * with an active auto-filter copies visible rows only.
 *
 * Univer's built-in clipboard hook already skips rows its own filter model
 * marks as filtered, which covers filters applied inside the app. Workbooks
 * opened from file are different: the sidecar delivers the filter range but
 * not its criteria, so the app re-applies the file's filtered-out rows as
 * plain hidden rows (`hideRows`). Those rows are invisible to the user but
 * not "filtered" to Univer, and the default copy would include them. This
 * hook reports raw-hidden rows inside the filter range as filtered-out too.
 * Hidden rows outside a filter range keep copying: manually hidden rows are
 * still included in copies.
 */
import { IUniverInstanceService, type IRange } from '@univerjs/core'
import { ISheetClipboardService } from '@univerjs/preset-sheets-core'
// The plugin package, not the preset: the preset re-exports the filter UI,
// whose module scope needs a browser canvas (Path2D) and cannot load in
// node-side tests.
import { SheetsFilterService } from '@univerjs/sheets-filter'

import type { UniverRuntime } from './univer-state'

/// Rows of `range` that fall inside `filterRange` and are hidden. Pure so the
/// row-exclusion rule is unit-testable without a Univer instance.
export function hiddenRowsInFilterRange(
  range: IRange,
  filterRange: IRange | null,
  isRowHidden: (row: number) => boolean,
): number[] {
  if (!filterRange) return []
  const start = Math.max(range.startRow, filterRange.startRow)
  const end = Math.min(range.endRow, filterRange.endRow)
  const rows: number[] = []
  for (let row = start; row <= end; row += 1) {
    if (isRowHidden(row)) rows.push(row)
  }
  return rows
}

export function installFilteredCopyHook(runtime: UniverRuntime): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const clipboardService = injector.get(ISheetClipboardService)
  const filterService = injector.get(SheetsFilterService)
  const instanceService = injector.get(IUniverInstanceService)
  return clipboardService.addClipboardHook({
    id: 'genoffice-filtered-copy',
    getFilteredOutRows: (unitId, subUnitId, range) => {
      const filterModel = filterService.getFilterModel(unitId, subUnitId)
      const worksheet = instanceService.getUniverSheetInstance(unitId)?.getSheetBySheetId(subUnitId)
      if (!filterModel || !worksheet) return []
      return hiddenRowsInFilterRange(
        range,
        filterModel.getRange(),
        // Raw visibility ignores the filter model — model-filtered rows are
        // already excluded by Univer's own default hook.
        (row) => !worksheet.getRowRawVisible(row),
      )
    },
  })
}
