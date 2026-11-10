/**
 * Excel's Ctrl+drag on the fill handle (Cmd on Mac): a lone number fills a
 * series (1008, 1009, …) instead of repeating, and a source that would fill
 * a series repeats instead. Univer's default hook picks COPY for a single
 * numeric cell and SERIES otherwise; an APPEND hook's answer outranks it,
 * so this one flips the choice whenever the modifier was down when the drag
 * ended (public issue genoffice#808).
 */
import {
  CellValueType,
  type ICellData,
  IUniverInstanceService,
  type Nullable,
  UniverInstanceType,
  type Workbook,
} from '@univerjs/core'
import {
  AUTO_FILL_APPLY_TYPE,
  AUTO_FILL_HOOK_TYPE,
  IAutoFillService,
  type IAutoFillLocation,
} from '@univerjs/sheets'

import type { UniverRuntime } from './univer-state'

export function invertedFillType(
  source: IAutoFillLocation['source'],
  cellAt: (row: number, column: number) => Nullable<ICellData>,
): AUTO_FILL_APPLY_TYPE {
  const [row, ...moreRows] = source.rows
  const [column, ...moreColumns] = source.cols
  if (row === undefined || column === undefined || moreRows.length || moreColumns.length) {
    return AUTO_FILL_APPLY_TYPE.COPY
  }
  const cell = cellAt(row, column)
  // a formula's cached result is numeric too; the formula itself repeats
  if (cell?.f || cell?.si) return AUTO_FILL_APPLY_TYPE.COPY
  const numeric = typeof cell?.v === 'number' || cell?.t === CellValueType.NUMBER
  return numeric ? AUTO_FILL_APPLY_TYPE.SERIES : AUTO_FILL_APPLY_TYPE.COPY
}

export function installCtrlDragFill(runtime: UniverRuntime): { dispose(): void } {
  // the fill handle's pointerup dispatches the auto-fill command; a capture
  // listener sees the same event first
  let modifier = false
  const onPointerDown = () => {
    modifier = false
  }
  const onPointerUp = (event: PointerEvent) => {
    modifier = event.ctrlKey || event.metaKey
  }
  window.addEventListener('pointerdown', onPointerDown, true)
  window.addEventListener('pointerup', onPointerUp, true)
  const injector = runtime.univer.__getInjector()
  const hook = injector.get(IAutoFillService).addHook({
    id: 'genoffice.ctrl-drag-fill',
    type: AUTO_FILL_HOOK_TYPE.APPEND,
    onBeforeFillData: (location) => {
      if (!modifier) return undefined
      const sheet = injector
        .get(IUniverInstanceService)
        .getUnit<Workbook>(location.unitId, UniverInstanceType.UNIVER_SHEET)
        ?.getSheetBySheetId(location.subUnitId)
      if (!sheet) return undefined
      return invertedFillType(location.source, (row, column) => sheet.getCellRaw(row, column))
    },
  })
  return {
    dispose() {
      hook.dispose()
      window.removeEventListener('pointerdown', onPointerDown, true)
      window.removeEventListener('pointerup', onPointerUp, true)
    },
  }
}
