import { RANGE_TYPE, type IRange } from '@univerjs/core'
import {
  DeltaColumnWidthCommand,
  SetColWidthCommand,
  SheetsSelectionsService,
} from '@univerjs/sheets'

import type { LazyWorkbookState, UniverRuntime } from './univer-state'

/**
 * Width lives on the column. Univer also uses the command's row bounds to
 * measure wrapped text and queue auto-height work. Bound that work to the
 * resident viewport on streamed sheets so it never queues the unloaded tail.
 */
export function installStreamedColumnWidth(
  runtime: UniverRuntime,
  workbookRef: { readonly current: LazyWorkbookState | null },
): { dispose(): void } {
  const selectionsService = runtime.univer.__getInjector().get(SheetsSelectionsService)
  return runtime.univerAPI.addEvent(runtime.univerAPI.Event.BeforeCommandExecute, (event) => {
    const drag = event.id === DeltaColumnWidthCommand.id
    if (event.cancel || (!drag && event.id !== SetColWidthCommand.id)) return
    const state = workbookRef.current
    if (!state || state.flags.preloadComplete) return
    const params = event.params as
      | {
          anchorCol?: number
          deltaX?: number
          unitId?: string
          subUnitId?: string
          ranges?: IRange[]
        }
      | undefined
    if (!params) return
    const workbook = runtime.univerAPI.getActiveWorkbook()
    if (!workbook || (params.unitId && params.unitId !== workbook.getId())) return
    const worksheet = params.subUnitId
      ? workbook.getSheetBySheetId(params.subUnitId)
      : workbook.getActiveSheet()
    if (!worksheet) return
    const sheetId = worksheet.getSheetId()
    if (!state.file.sheets.some((sheet) => sheet.id === sheetId)) return

    const selections = selectionsService.getCurrentSelections().map(({ range }) => range)
    const resident = state.loadedRanges.get(sheetId)
    const visible = worksheet.getVisibleRange()
    const startRow = Math.max(resident?.startRow ?? 0, visible?.startRow ?? 0)
    const endRow = Math.max(startRow, Math.min(resident?.endRow ?? 0, visible?.endRow ?? startRow))
    const boundRows = ({ startColumn, endColumn, rangeType }: IRange): IRange => ({
      startColumn: rangeType === RANGE_TYPE.ROW || rangeType === RANGE_TYPE.ALL ? 0 : startColumn,
      endColumn:
        rangeType === RANGE_TYPE.ROW || rangeType === RANGE_TYPE.ALL
          ? worksheet.getMaxColumns() - 1
          : endColumn,
      startRow,
      endRow,
    })
    if (!drag) {
      const ranges = params.ranges?.length ? params.ranges : selections
      params.ranges = ranges.map(boundRows)
      return
    }

    const { anchorCol, deltaX } = params
    if (anchorCol === undefined || deltaX === undefined || selections.length === 0) return
    const columns = selections.filter((range) => range.rangeType === RANGE_TYPE.COLUMN)
    const ranges =
      selections.length === 1 && selections[0]?.rangeType === RANGE_TYPE.ALL
        ? selections.map(boundRows)
        : columns.some(
              ({ startColumn, endColumn }) => anchorCol >= startColumn && anchorCol <= endColumn,
            )
          ? columns.map(boundRows)
          : [{ startColumn: anchorCol, endColumn: anchorCol, startRow, endRow }]
    const value = worksheet.getColumnWidth(anchorCol) + deltaX
    event.cancel = true
    queueMicrotask(() => {
      if (workbookRef.current !== state) return
      // The native width command owns the mutation, protection check and
      // undo/redo. Its mutation is also recorded by the XLSX save journal.
      void runtime.univerAPI.executeCommand(SetColWidthCommand.id, {
        unitId: workbook.getId(),
        subUnitId: sheetId,
        ranges,
        value,
      })
    })
  })
}
