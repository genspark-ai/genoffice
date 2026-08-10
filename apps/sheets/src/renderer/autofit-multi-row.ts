/**
 * Excel-style multi-row AutoFit. Univer's row-header double-click resize
 * autofits only the double-clicked row (the column counterpart already walks
 * the whole selection — upstream asymmetry). When the clicked row lies
 * inside a multi-row full-row selection, the single-row command is cancelled
 * and re-dispatched with every selected row span, so each selected row sizes
 * to its own content, matching Excel.
 */
import type { IRange } from '@univerjs/core'

import type { UniverRuntime } from './univer-state'

export const SET_ROW_IS_AUTO_HEIGHT_COMMAND = 'sheet.command.set-worksheet-row-is-auto-height'

/// Full-row spans: selections made from the row headers or a select-all.
function isFullRowSpan(range: IRange, columnCount: number): boolean {
  return range.startColumn === 0 && range.endColumn >= columnCount - 1
}

/**
 * Returns the replacement ranges when the command should grow to the whole
 * selection, or null to leave the command untouched. Grows only when the
 * command targets a single row that sits inside a full-row selection
 * covering more than one row in total (Ctrl-selected disjoint spans count).
 */
export function expandAutoHeightRanges(
  commandRanges: readonly IRange[],
  selections: readonly IRange[],
  columnCount: number,
): IRange[] | null {
  if (commandRanges.length !== 1) return null
  const clicked = commandRanges[0]
  if (clicked === undefined || clicked.startRow !== clicked.endRow) return null
  const rowSpans = selections.filter((range) => isFullRowSpan(range, columnCount))
  if (
    !rowSpans.some(
      (range) => range.startRow <= clicked.startRow && clicked.startRow <= range.endRow,
    )
  ) {
    return null
  }
  const totalRows = rowSpans.reduce((sum, range) => sum + (range.endRow - range.startRow + 1), 0)
  if (totalRows <= 1) return null
  return rowSpans.map((range) => ({
    startRow: range.startRow,
    endRow: range.endRow,
    startColumn: clicked.startColumn,
    endColumn: clicked.endColumn,
  }))
}

export function installMultiRowAutofit(runtime: UniverRuntime): { dispose(): void } {
  let redispatching = false
  return runtime.univerAPI.addEvent(runtime.univerAPI.Event.BeforeCommandExecute, (event) => {
    if (redispatching || event.id !== SET_ROW_IS_AUTO_HEIGHT_COMMAND) return
    const params = (event.params ?? {}) as { ranges?: IRange[] }
    if (!params.ranges) return
    const worksheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
    if (!worksheet) return
    const selections = (worksheet.getSelection()?.getActiveRangeList() ?? []).map((range) =>
      range.getRange(),
    )
    const expanded = expandAutoHeightRanges(
      params.ranges,
      selections,
      worksheet.getSheet().getColumnCount(),
    )
    if (expanded === null) return
    event.cancel = true
    redispatching = true
    void runtime.univerAPI
      .executeCommand(SET_ROW_IS_AUTO_HEIGHT_COMMAND, { ...event.params, ranges: expanded })
      .finally(() => {
        redispatching = false
      })
  })
}
