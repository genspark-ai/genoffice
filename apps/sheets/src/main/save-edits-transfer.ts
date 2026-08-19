import type {
  WorkbookCellEdit,
  WorkbookSaveEditsBegin,
  WorkbookSaveEditsChunk,
} from '../shared/desktop-api'
import { MAX_SAVE_EDITS_TOTAL } from '../shared/ipc-channels'

/// A tab keeps at most this many transfers open at once. A save consumes its
/// transfer immediately, so anything beyond "one save plus one recovery tick
/// racing" indicates a renderer bug or an abandoned upload.
const MAX_OPEN_TRANSFERS = 4

/// An unfinished upload this old is abandoned (the renderer retries a failed
/// save from scratch with a fresh transfer id).
const TRANSFER_IDLE_EXPIRY_MS = 120_000

interface PendingTransfer {
  readonly sessionId: string
  readonly total: number
  readonly chunks: (readonly WorkbookCellEdit[])[]
  received: number
  lastActivityMs: number
}

/// Accumulates the ordered edit chunks of large saves, one store per sheets
/// tab. Chunking bounds each IPC message's structured-clone spike; the store
/// re-interns repeated sheetId strings so ten million deserialized edits do
/// not carry ten million copies of the same id.
export class SaveEditsTransferStore {
  private readonly transfers = new Map<string, PendingTransfer>()
  /// Runs while any transfer is open, so an abandoned upload frees its edits
  /// after the idle expiry even if no further transfer activity ever arrives.
  private sweepTimer: ReturnType<typeof setInterval> | null = null

  begin(request: WorkbookSaveEditsBegin, nowMs = Date.now()): void {
    this.sweep(nowMs)
    if (this.transfers.has(request.transferId)) {
      throw new Error('Save transfer already exists.')
    }
    if (this.transfers.size >= MAX_OPEN_TRANSFERS) {
      throw new Error('Too many open save transfers.')
    }
    this.transfers.set(request.transferId, {
      sessionId: request.sessionId,
      total: request.total,
      chunks: [],
      received: 0,
      lastActivityMs: nowMs,
    })
    this.syncSweepTimer()
  }

  addChunk(request: WorkbookSaveEditsChunk, nowMs = Date.now()): void {
    this.sweep(nowMs)
    const transfer = this.transfers.get(request.transferId)
    if (!transfer || transfer.sessionId !== request.sessionId) {
      throw new Error('Unknown save transfer.')
    }
    if (request.seq !== transfer.chunks.length) {
      this.discard(request.transferId, request.sessionId)
      throw new Error('Save transfer chunk out of order.')
    }
    if (
      transfer.received + request.edits.length > transfer.total ||
      transfer.received + request.edits.length > MAX_SAVE_EDITS_TOTAL
    ) {
      this.discard(request.transferId, request.sessionId)
      throw new Error('Save transfer exceeds its declared size.')
    }
    transfer.chunks.push(internSheetIds(request.edits))
    transfer.received += request.edits.length
    transfer.lastActivityMs = nowMs
  }

  /// Consumes the transfer: a transfer feeds exactly one save request.
  take(transferId: string, sessionId: string): WorkbookCellEdit[] {
    const transfer = this.transfers.get(transferId)
    if (!transfer || transfer.sessionId !== sessionId) {
      throw new Error('Unknown save transfer.')
    }
    this.discard(transferId, sessionId)
    if (transfer.received !== transfer.total) {
      throw new Error('Save transfer is incomplete.')
    }
    const edits: WorkbookCellEdit[] = []
    for (const chunk of transfer.chunks) {
      for (const edit of chunk) edits.push(edit)
    }
    return edits
  }

  /// Silent no-op when the transfer is already consumed, expired, or owned by
  /// another session: the renderer aborts best-effort from failure paths.
  discard(transferId: string, sessionId: string): void {
    const transfer = this.transfers.get(transferId)
    if (transfer && transfer.sessionId === sessionId) this.transfers.delete(transferId)
    this.syncSweepTimer()
  }

  discardSession(sessionId: string): void {
    for (const [transferId, transfer] of this.transfers) {
      if (transfer.sessionId === sessionId) this.transfers.delete(transferId)
    }
    this.syncSweepTimer()
  }

  sweep(nowMs = Date.now(), idleExpiryMs = TRANSFER_IDLE_EXPIRY_MS): void {
    for (const [transferId, transfer] of this.transfers) {
      if (nowMs - transfer.lastActivityMs > idleExpiryMs) this.transfers.delete(transferId)
    }
    this.syncSweepTimer()
  }

  /// Tab teardown: drop everything now — no reason to keep megabytes of
  /// edits reachable through the sweep timer until the idle expiry fires.
  dispose(): void {
    this.transfers.clear()
    this.syncSweepTimer()
  }

  get openTransferCount(): number {
    return this.transfers.size
  }

  private syncSweepTimer(): void {
    if (this.transfers.size > 0 && this.sweepTimer === null) {
      this.sweepTimer = setInterval(() => this.sweep(), TRANSFER_IDLE_EXPIRY_MS / 4)
      // Node timers must not keep the process alive for this
      this.sweepTimer.unref?.()
    } else if (this.transfers.size === 0 && this.sweepTimer !== null) {
      clearInterval(this.sweepTimer)
      this.sweepTimer = null
    }
  }
}

/// Structured clone gives every edit its own copy of the sheet id string;
/// across millions of edits that dwarfs the payload itself. Point them all at
/// one instance per sheet.
function internSheetIds(edits: readonly WorkbookCellEdit[]): readonly WorkbookCellEdit[] {
  const interned = new Map<string, string>()
  for (const edit of edits) {
    const existing = interned.get(edit.sheetId)
    if (existing === undefined) {
      interned.set(edit.sheetId, edit.sheetId)
    } else {
      ;(edit as { sheetId: string }).sheetId = existing
    }
  }
  return edits
}
