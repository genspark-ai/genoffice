/**
 * Copy materializes its selection first. On streamed workbooks the
 * clipboard serializes whatever the lazy loader has put in the cell matrix:
 * rows the viewport never passed lose their text constants entirely and
 * numbers lose their number format. Loading the selection into the lazy
 * window before Univer's copy runs makes one-shot big-range copies match
 * what the screen would show.
 */
import { ISheetClipboardService } from '@univerjs/sheets-ui'

import { formatAddress } from '../domain/cell-address'
import { t } from './i18n/locale'
import { ensureLazyRangeLoaded } from './univer-sync'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

/// readWorkbookRange's protocol cap (MAX_RANGE_CELLS, desktop-api.ts); the
/// lazy window holds one contiguous range, so a selection past this cannot
/// be materialized in full — warn instead of silently copying blanks.
const MATERIALIZE_CELL_LIMIT = 20_000

async function materializeSelection(
  runtime: UniverRuntime,
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null },
  setMessage: (message: string) => void,
): Promise<void> {
  const state = lazyWorkbookRef.current
  if (!state || state.formulaMode) return
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const selection = workbook?.getActiveRange()
  if (!workbook || !worksheet || !selection) return
  const sheetMeta = state.file.sheets.find((sheet) => sheet.id === worksheet.getSheetId())
  if (!sheetMeta) return
  const range = {
    startRow: selection.getRow(),
    startColumn: selection.getColumn(),
    endRow: Math.min(selection.getRow() + selection.getHeight() - 1, sheetMeta.rowCount - 1),
    endColumn: Math.min(
      selection.getColumn() + selection.getWidth() - 1,
      sheetMeta.columnCount - 1,
    ),
  }
  const loaded = state.loadedRanges.get(worksheet.getSheetId())
  if (
    loaded &&
    range.startRow >= loaded.startRow &&
    range.endRow <= loaded.endRow &&
    range.startColumn >= loaded.startColumn &&
    range.endColumn <= loaded.endColumn
  ) {
    return
  }
  const cells = (range.endRow - range.startRow + 1) * (range.endColumn - range.startColumn + 1)
  if (cells > MATERIALIZE_CELL_LIMIT) {
    const rangeText = `${formatAddress(range.startRow, range.startColumn)}:${formatAddress(range.endRow, range.endColumn)}`
    setMessage(t('appRangeTooManyCells', { range: rangeText, max: MATERIALIZE_CELL_LIMIT }))
    return
  }
  try {
    await ensureLazyRangeLoaded(runtime, lazyWorkbookRef, worksheet, range, setMessage)
  } catch {
    // Copy still runs on whatever is materialized; same as before the fix.
  }
}

export function installCopyMaterialize(
  runtime: UniverRuntime,
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null },
  setMessage: (message: string) => void,
): { dispose(): void } {
  const clipboardService = runtime.univer.__getInjector().get(ISheetClipboardService) as {
    copy(options?: unknown): Promise<boolean>
    cut(): Promise<boolean>
  }
  const originalCopy = clipboardService.copy.bind(clipboardService)
  const originalCut = clipboardService.cut.bind(clipboardService)
  clipboardService.copy = async (options?: unknown) => {
    await materializeSelection(runtime, lazyWorkbookRef, setMessage)
    return originalCopy(options)
  }
  clipboardService.cut = async () => {
    await materializeSelection(runtime, lazyWorkbookRef, setMessage)
    return originalCut()
  }
  return {
    dispose() {
      clipboardService.copy = originalCopy
      clipboardService.cut = originalCut
    },
  }
}
