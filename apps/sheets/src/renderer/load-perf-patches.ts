/**
 * Univer rebuilds every header "unhide" arrow (one shape per hidden run, each
 * disposal scheduling its own MessageChannel task) on every skeleton change.
 * A streamed load of a filtered sheet with a thousand hidden runs therefore
 * paid a full rebuild per chunk. Coalesce the rebuilds so a burst of skeleton
 * changes ends in one.
 */
import { IUniverInstanceService } from '@univerjs/core'
import { SheetRangeThemeModel } from '@univerjs/sheets'
import type { UniverRuntime } from './univer-state'
import { HeaderUnhideRenderController } from '@univerjs/sheets-ui'

const HEADER_UNHIDE_SETTLE_MS = 120

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
