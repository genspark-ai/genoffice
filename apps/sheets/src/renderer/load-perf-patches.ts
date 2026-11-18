/**
 * Univer rebuilds every header "unhide" arrow (one shape per hidden run, each
 * disposal scheduling its own MessageChannel task) on every skeleton change.
 * A streamed load of a filtered sheet with a thousand hidden runs therefore
 * paid a full rebuild per chunk. Coalesce the rebuilds so a burst of skeleton
 * changes ends in one.
 */
import {
  IUniverInstanceService,
  SheetSkeleton,
  type Injector,
  type IRowData,
  type Worksheet,
} from '@univerjs/core'
import { INTERCEPTOR_POINT, SheetInterceptorService, SheetRangeThemeModel } from '@univerjs/sheets'
import { SheetsFilterService } from '@univerjs/sheets-filter'
import { TableManager } from '@univerjs/sheets-table'
import type { UniverRuntime } from './univer-state'
import { HeaderUnhideRenderController } from '@univerjs/sheets-ui'

const HEADER_UNHIDE_SETTLE_MS = 120
const LARGE_ROW_COUNT = 100_000

interface RowGeometry {
  rowTotalHeight: number
  rowHeightAccumulation: number[]
}

interface RowLayout {
  readonly worksheet: Worksheet
  readonly _injector: Injector
  readonly gapConfig: { rowGaps?: object }
  _generateRowMatrixCache(
    rowCount: number,
    rows: Record<number, IRowData | undefined>,
    defaultHeight: number,
  ): RowGeometry
}

let rowGeometryInstalled = false

/** Univer's row-filter interceptors are only needed when there is a live
 * sheet filter or table. Preserve its accumulated-height array for scroll and
 * hit testing, and handle sparse row overrides outside the hot loop. */
export function installSparseRowGeometry(): void {
  if (rowGeometryInstalled) return
  const prototype = SheetSkeleton.prototype as unknown as RowLayout
  const generate = prototype._generateRowMatrixCache
  if (typeof generate !== 'function') return
  rowGeometryInstalled = true
  const uniform = new WeakMap<
    RowLayout,
    { rowCount: number; defaultHeight: number; geometry: RowGeometry }
  >()
  prototype._generateRowMatrixCache = function (rowCount, rows, defaultHeight) {
    if (rowCount < LARGE_ROW_COUNT || this.gapConfig.rowGaps || hasRowFiltering(this)) {
      return generate.call(this, rowCount, rows, defaultHeight)
    }

    const overrides = Object.keys(rows)
      .map(Number)
      .filter((row) => row >= 0 && row < rowCount && rows[row] != null)
    if (overrides.length === 0) {
      const cached = uniform.get(this)
      if (cached?.rowCount === rowCount && cached.defaultHeight === defaultHeight) {
        return cached.geometry
      }
    }
    const rowHeightAccumulation = new Array<number>(rowCount)
    let rowTotalHeight = 0
    let row = 0
    for (const at of overrides) {
      for (; row < at; row += 1) {
        rowTotalHeight += defaultHeight
        rowHeightAccumulation[row] = rowTotalHeight
      }
      const data = rows[at]!
      const height =
        (data.ia == null || data.ia === 1) && typeof data.ah === 'number' && data.ah > 0
          ? data.ah
          : (data.h ?? defaultHeight)
      rowTotalHeight += data.hd === 1 ? 0 : height
      rowHeightAccumulation[row++] = rowTotalHeight
    }
    for (; row < rowCount; row += 1) {
      rowTotalHeight += defaultHeight
      rowHeightAccumulation[row] = rowTotalHeight
    }
    const geometry = { rowTotalHeight, rowHeightAccumulation }
    if (overrides.length === 0) uniform.set(this, { rowCount, defaultHeight, geometry })
    return geometry
  }
}

function hasRowFiltering({ worksheet, _injector }: RowLayout): boolean {
  const { _interceptorsByName } = _injector.get(SheetInterceptorService) as unknown as {
    _interceptorsByName?: Map<unknown, readonly { priority?: number }[]>
  }
  // Fall back for a different Univer implementation or added interceptors.
  const handlers = _interceptorsByName?.get(INTERCEPTOR_POINT.ROW_FILTERED)
  if (
    handlers?.length !== 2 ||
    handlers[0]?.priority !== 100 ||
    (handlers[1]?.priority ?? 0) !== 0
  ) {
    return true
  }
  const unitId = worksheet.getUnitId()
  const sheetId = worksheet.getSheetId()
  return (
    _injector.get(SheetsFilterService).getFilterModel(unitId, sheetId) != null ||
    _injector.get(TableManager).getTablesBySubunitId(unitId, sheetId).length !== 0
  )
}

let installed = false

export function installHeaderUnhideDebounce(): void {
  if (installed) return
  installed = true
  const proto = HeaderUnhideRenderController.prototype as any
  const origUpdate = proto._update
  const pending = new WeakMap<object, ReturnType<typeof setTimeout>>()
  proto._update = function (workbook: unknown, worksheet: unknown) {
    const timer = pending.get(this)
    if (timer !== undefined) clearTimeout(timer)
    pending.set(
      this,
      setTimeout(() => {
        pending.delete(this)
        if (this._disposed) return
        origUpdate.call(this, workbook, worksheet)
      }, HEADER_UNHIDE_SETTLE_MS),
    )
  }
}

/** Alternating range styles are the only consumers of the full-sheet row
 * visibility cache. On large sheets without range styles, leave the cache
 * uninitialized so registering the first rule still builds it on demand. */
export function installRangeThemeVisibilityGate(runtime: UniverRuntime): void {
  const injector = runtime.univer.__getInjector()
  const instances = injector.get(IUniverInstanceService)
  const themes = injector.get(SheetRangeThemeModel)
  const original = themes.refreshSheetRowVisibleFuncSet
  const cache = (
    themes as unknown as {
      _rowVisibleFuncSet?: Map<string, Map<string, Set<number>>>
    }
  )._rowVisibleFuncSet
  if (!cache) return
  themes.refreshSheetRowVisibleFuncSet = (unitId, subUnitId) => {
    const sheet = instances.getUniverSheetInstance(unitId)?.getSheetBySheetId(subUnitId)
    if (
      sheet &&
      sheet.getRowCount() >= 100_000 &&
      !themes.getRegisteredRangeThemeStyle({
        unitId,
        subUnitId,
        range: {
          startRow: 0,
          endRow: sheet.getRowCount() - 1,
          startColumn: 0,
          endColumn: sheet.getColumnCount() - 1,
        },
      })
    ) {
      cache.get(unitId)?.delete(subUnitId)
      return
    }
    original.call(themes, unitId, subUnitId)
  }
}
