import { columnIndex } from '@genoffice/xlsx-gateway/domain/cell-address'
import type { SheetVisual } from '@genoffice/xlsx-gateway/domain/chart-visual'
import type { WorkbookSnapshot } from '@genoffice/xlsx-gateway/domain/workbook.types'

import { applyFormatPatchToRange, syncUniver } from './univer-sync'
import { collabRemoteApply, type UniverRuntime } from './univer-state'

/**
 * Apply a snapshot that arrived from another member.
 *
 * Values, formulas, formats, merges and row/column sizes are replayed in place
 * so the unit — and its undo history — survive. The writes are deliberately
 * NOT suppressed from the edit journal: the journal is the save source, so a
 * peer's edits must be in it or the next local save would silently drop them.
 * `collabRemoteApply` keeps them out of the undo stack; App suppresses the
 * publish echo while it is set.
 *
 * Visuals are not Univer cells, so they cannot ride the command listener. They
 * are handed back to the caller, which merges them into `editJournal.visualAdds`
 * (the save source and the visual installer's input) and reinstalls floats.
 */
export interface RemoteApplyResult {
  /** charts the peer has on each sheet; callers merge these into the journal */
  readonly visuals: readonly SheetVisual[]
}

export function applyRemoteWorkbookSnapshot(
  runtime: UniverRuntime | null,
  raw: string,
): RemoteApplyResult {
  let snapshot: WorkbookSnapshot
  try {
    snapshot = JSON.parse(raw) as WorkbookSnapshot
  } catch {
    return { visuals: [] }
  }
  const visuals = snapshot.sheets.flatMap((sheet) => sheet.visuals ?? [])
  collabRemoteApply.active = true
  try {
    syncUniver(runtime, snapshot)
    const workbook = runtime?.univerAPI.getActiveWorkbook()
    if (workbook) {
      for (const sheet of snapshot.sheets) {
        const worksheet = workbook.getSheetBySheetId(sheet.id)
        if (!worksheet) continue
        for (const [address, format] of Object.entries(sheet.styles ?? {})) {
          applyFormatPatchToRange(worksheet.getRange(address), format)
        }
        for (const merge of sheet.merges ?? []) {
          worksheet.getRange(merge).merge()
        }
        for (const [row, heightPoints] of Object.entries(sheet.rowHeights ?? {})) {
          worksheet.setRowHeights(Number(row) - 1, 1, Math.round((heightPoints * 96) / 72))
        }
        for (const [column, widthPx] of Object.entries(sheet.colWidths ?? {})) {
          worksheet.setColumnWidths(columnIndex(column), 1, Math.round(widthPx))
        }
      }
    }
  } finally {
    collabRemoteApply.active = false
  }
  return { visuals }
}
