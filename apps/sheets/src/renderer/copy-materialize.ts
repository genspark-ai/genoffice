/**
 * Copy materializes its selection first. On streamed workbooks the
 * clipboard serializes whatever the lazy loader has put in the cell matrix:
 * rows the viewport never passed lose their text constants entirely and
 * numbers lose their number format. Loading the selection into the lazy
 * window before Univer's copy runs makes one-shot big-range copies match
 * what the screen would show.
 *
 * Past the full-load budget the selection cannot be made resident at all:
 * the plain-text slice is built straight from the sidecar payload instead
 * (values as displayed, no styles) so the clipboard never carries blanks
 * for cells the file has.
 */
import { CellValueType, numfmt } from '@univerjs/core'
import { ISheetClipboardService } from '@univerjs/sheets-ui'

import { formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import { FULL_LOAD_MAX_CELLS } from './app-constants'
import { clipboardField } from './clipboard-tsv'
import { t } from './i18n/locale'
import { ensureLazyRangeLoaded, readCopySourceDirect, type RawCopyCell } from './univer-sync'
import {
  lazySheetMeta,
  lazySheetScreenExtent,
  type LazyWorkbookState,
  type UniverRuntime,
} from './univer-state'

/// A selection this size loads as one resident window (the same budget the
/// full-load offer uses); Univer then copies styles and all.
export const COPY_MATERIALIZE_MAX_CELLS = FULL_LOAD_MAX_CELLS

type CopyOutcome = 'grid' | 'done' | 'refused'

/// TSV text for a directly read cell: number formats apply like the grid
/// would show them; rich text collapsed to its plain run text upstream.
export function directCopyField(cell: RawCopyCell | undefined): string {
  if (!cell || cell.v === null || cell.v === undefined) return ''
  const pattern = (cell.s as { n?: { pattern?: string } } | null)?.n?.pattern
  if (typeof cell.v === 'number' && pattern && pattern !== 'General') {
    const text = numfmt.format(pattern, cell.v, { throws: false })
    if (typeof text === 'string') return clipboardField({ v: text, t: CellValueType.STRING })
  }
  return clipboardField({
    v: cell.v,
    t: cell.t === CellValueType.BOOLEAN ? CellValueType.BOOLEAN : undefined,
  })
}

export function directCopyPlainText(rows: readonly (readonly RawCopyCell[])[]): string {
  return rows.map((row) => row.map((cell) => directCopyField(cell)).join('\t')).join('\n')
}

async function materializeSelection(
  runtime: UniverRuntime,
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null },
  setMessage: (message: string) => void,
  kind: 'copy' | 'cut',
): Promise<CopyOutcome> {
  const state = lazyWorkbookRef.current
  if (!state || state.formulaMode) return 'grid'
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  const selection = workbook?.getActiveRange()
  if (!workbook || !worksheet || !selection) return 'grid'
  const sheetMeta = lazySheetMeta(state, worksheet.getSheetId())
  if (!sheetMeta) return 'grid'
  // Selection coordinates are screen-space, so the clamp must be the screen
  // extent (file extent shifted by this session's structural ops), not the
  // file extent: after insert/delete ops a file-space clamp cuts off or
  // overshoots the last rows/columns (cf. applyRangeInLoadedChunks in
  // univer-sync.ts, which documents the same rule).
  const extent = lazySheetScreenExtent(state, worksheet.getSheetId()) ?? {
    rows: sheetMeta.rowCount,
    columns: sheetMeta.columnCount,
  }
  const range = {
    startRow: selection.getRow(),
    startColumn: selection.getColumn(),
    endRow: Math.min(selection.getRow() + selection.getHeight() - 1, extent.rows - 1),
    endColumn: Math.min(selection.getColumn() + selection.getWidth() - 1, extent.columns - 1),
  }
  // A selection wholly past the extent (stale anchor after deletes) inverts
  // the range: nothing to load, and downstream math must not see it.
  if (range.endRow < range.startRow || range.endColumn < range.startColumn) return 'grid'
  const loaded = state.loadedRanges.get(worksheet.getSheetId())
  if (
    loaded &&
    range.startRow >= loaded.startRow &&
    range.endRow <= loaded.endRow &&
    range.startColumn >= loaded.startColumn &&
    range.endColumn <= loaded.endColumn
  ) {
    return 'grid'
  }
  const cells = (range.endRow - range.startRow + 1) * (range.endColumn - range.startColumn + 1)
  const rangeText = `${formatAddress(range.startRow, range.startColumn)}:${formatAddress(range.endRow, range.endColumn)}`
  if (cells > COPY_MATERIALIZE_MAX_CELLS) {
    // A cut past the resident budget would clear cells the clipboard never
    // received; only copy has the values-only path.
    if (kind === 'cut') {
      setMessage(t('appRangeTooManyCells', { range: rangeText, max: COPY_MATERIALIZE_MAX_CELLS }))
      return 'refused'
    }
    setMessage(t('appCopyLoadingRange', { range: rangeText }))
    const source = await readCopySourceDirect(
      lazyWorkbookRef,
      worksheet.getSheetId(),
      range,
      setMessage,
    )
    if (!source || lazyWorkbookRef.current !== state) {
      setMessage(t('appRangeTooManyCells', { range: rangeText, max: COPY_MATERIALIZE_MAX_CELLS }))
      return 'refused'
    }
    try {
      await navigator.clipboard.writeText(directCopyPlainText(source))
    } catch {
      setMessage(t('appRangeTooManyCells', { range: rangeText, max: COPY_MATERIALIZE_MAX_CELLS }))
      return 'refused'
    }
    setMessage(
      t('appCopyValuesOnly', {
        range: rangeText,
        cells: cells.toLocaleString(),
        max: COPY_MATERIALIZE_MAX_CELLS.toLocaleString(),
      }),
    )
    return 'done'
  }
  if (cells > 20_000) setMessage(t('appCopyLoadingRange', { range: rangeText }))
  try {
    await ensureLazyRangeLoaded(runtime, lazyWorkbookRef, worksheet, range, setMessage)
  } catch {
    // Copy still runs on whatever is materialized; same as before the fix.
  }
  return 'grid'
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
    const outcome = await materializeSelection(runtime, lazyWorkbookRef, setMessage, 'copy')
    return outcome === 'grid' ? originalCopy(options) : outcome === 'done'
  }
  clipboardService.cut = async () => {
    const outcome = await materializeSelection(runtime, lazyWorkbookRef, setMessage, 'cut')
    return outcome === 'grid' ? originalCut() : false
  }
  return {
    dispose() {
      clipboardService.copy = originalCopy
      clipboardService.cut = originalCut
    },
  }
}
