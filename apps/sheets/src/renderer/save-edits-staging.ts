import type { WorkbookCellEdit } from '../shared/desktop-api'
import {
  MAX_SAVE_EDITS_TOTAL,
  SAVE_EDITS_CHUNK_MAX,
  SAVE_EDITS_INLINE_MAX,
} from '../shared/ipc-channels'

/// The slice of desktopApi the staging step needs (injectable for tests).
export interface SaveEditsTransferApi {
  beginSaveEditsTransfer(request: {
    sessionId: string
    transferId: string
    total: number
  }): Promise<void>
  sendSaveEditsChunk(request: {
    sessionId: string
    transferId: string
    seq: number
    editsJson: string
  }): Promise<void>
  abortSaveEditsTransfer(request: { sessionId: string; transferId: string }): Promise<void>
}

export interface StagedEdits {
  edits: WorkbookCellEdit[]
  editsTransferId?: string
}

/**
 * Prepares a save's cell edits for the IPC hop. Small edit sets ride inline
 * in the save request as before; sets above SAVE_EDITS_INLINE_MAX are
 * uploaded to the main process as JSON-string chunks first (a flat string
 * crosses the context bridge in milliseconds where the same edits as a live
 * array cost microseconds per edit), and the request references the finished
 * transfer instead.
 */
export async function stageEditsForSave(
  api: SaveEditsTransferApi,
  sessionId: string,
  edits: WorkbookCellEdit[],
): Promise<StagedEdits> {
  if (edits.length <= SAVE_EDITS_INLINE_MAX) return { edits }
  if (edits.length > MAX_SAVE_EDITS_TOTAL) {
    throw new Error(
      `This save contains ${edits.length.toLocaleString()} cell edits, above the ` +
        `${MAX_SAVE_EDITS_TOTAL.toLocaleString()} maximum for a single save.`,
    )
  }
  const transferId = crypto.randomUUID()
  await api.beginSaveEditsTransfer({ sessionId, transferId, total: edits.length })
  try {
    for (let start = 0, seq = 0; start < edits.length; start += SAVE_EDITS_CHUNK_MAX, seq += 1) {
      await api.sendSaveEditsChunk({
        sessionId,
        transferId,
        seq,
        editsJson: JSON.stringify(edits.slice(start, start + SAVE_EDITS_CHUNK_MAX)),
      })
    }
  } catch (error) {
    // Free the partial upload's main-process memory now instead of leaving
    // it to the idle expiry (and toward the open-transfer cap).
    await abortStagedEditsTransfer(api, sessionId, transferId)
    throw error
  }
  return { edits: [], editsTransferId: transferId }
}

/// Best-effort: the transfer also expires on its own if this call fails.
export async function abortStagedEditsTransfer(
  api: SaveEditsTransferApi,
  sessionId: string,
  transferId: string | undefined,
): Promise<void> {
  if (transferId === undefined) return
  await api.abortSaveEditsTransfer({ sessionId, transferId }).catch(() => undefined)
}
