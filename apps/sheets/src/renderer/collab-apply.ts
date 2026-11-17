import type { WorkbookSnapshot } from '@genoffice/xlsx-gateway/domain/workbook.types'
import type { UniverRuntime } from './univer-state'
import { applyFormatPatchToRange, syncUniver } from './univer-sync'

/**
 * Apply a snapshot that arrived from another member.
 *
 * The values and the formulas go through the same path a file open uses, and
 * the formats are replayed per cell. Merges, row heights, column widths and
 * the AI visuals are not replayed yet — they ride in the snapshot but only
 * land on a full rebuild.
 */
export function applyRemoteWorkbookSnapshot(runtime: UniverRuntime | null, raw: string): void {
  let snapshot: WorkbookSnapshot
  try {
    snapshot = JSON.parse(raw) as WorkbookSnapshot
  } catch {
    return
  }
  syncUniver(runtime, snapshot)
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  if (!workbook) return
  for (const sheet of snapshot.sheets) {
    const worksheet = workbook.getSheetBySheetId(sheet.id)
    if (!worksheet) continue
    for (const [address, format] of Object.entries(sheet.styles ?? {})) {
      applyFormatPatchToRange(worksheet.getRange(address), format)
    }
  }
}
