import { MAIL_CHANNELS } from '../../shared/ipc'
import {
  mailOfflineStatusSchema,
  mailOutboxCancelSchema,
  mailOutboxListSchema,
  mailOutboxRetrySchema,
  mailSyncNowSchema,
} from '../../shared/ipc-schemas'
import type { SyncEngine, OfflineStatusPayload } from './engine'

/**
 * Sync/outbox IPC handlers (P4 spec section 6). Every invoke payload is
 * zod-validated end-to-end before touching the engine; the preload stays
 * zod-free with hand-rolled guards (sheets precedent). Main-process only.
 */

/** Minimal handler table so tests can inject a fake bus (ipcMain-compatible). */
export interface SyncIpcBus {
  handle(channel: string, listener: (payload: unknown) => Promise<unknown>): void
}

export interface SyncIpcDeps {
  /** Broadcast an offline-status event to attached renderers. */
  sendOfflineStatus?: (status: OfflineStatusPayload) => void
}

function badRequest(message: string): Error {
  return new Error(`mail: ${message}`)
}

export interface SyncIpcHandle {
  /** Subscribe to offline status; replays the current value on subscribe. */
  onOfflineStatus(handler: (status: OfflineStatusPayload) => void): () => void
}

/**
 * Register the four invoke channels and fan engine offline transitions
 * out to renderers. Offline status itself is an event channel (no invoke
 * handler); new subscribers get the current value replayed immediately.
 */
export function registerSyncHandlers(
  bus: SyncIpcBus,
  engine: SyncEngine,
  deps?: SyncIpcDeps,
): SyncIpcHandle {
  const watchers = new Set<(status: OfflineStatusPayload) => void>()
  const stopForward = engine.watchOffline((status) => {
    deps?.sendOfflineStatus?.(status)
    for (const watcher of watchers) watcher(status)
  })
  void stopForward

  bus.handle(MAIL_CHANNELS.syncNow, async (payload) => {
    const parsed = mailSyncNowSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad sync request')
    return engine.syncNow(parsed.data.accountId, parsed.data.folderId)
  })

  bus.handle(MAIL_CHANNELS.outboxList, async (payload) => {
    const parsed = mailOutboxListSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad outbox list request')
    return engine.outboxList(parsed.data.accountId)
  })

  bus.handle(MAIL_CHANNELS.outboxRetry, async (payload) => {
    const parsed = mailOutboxRetrySchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad outbox retry request')
    try {
      return engine.retryOutbox(parsed.data.accountId, parsed.data.outboxId)
    } catch {
      throw badRequest('unknown outbox entry')
    }
  })

  bus.handle(MAIL_CHANNELS.outboxCancel, async (payload) => {
    const parsed = mailOutboxCancelSchema.safeParse(payload)
    if (!parsed.success) throw badRequest('bad outbox cancel request')
    return engine.cancelOutbox(parsed.data.accountId, parsed.data.outboxId)
  })

  // Offline status is an event channel (no invoke handler): engine
  // transitions fan out through the watcher above, and new subscribers
  // get the current value replayed immediately.
  const handle: SyncIpcHandle = {
    onOfflineStatus: (handler) => {
      watchers.add(handler)
      // Replay the current value so late subscribers see live state.
      // The replay is schema-checked like every live event.
      const current = engine.currentOfflineStatus()
      const parsed = mailOfflineStatusSchema.safeParse(current)
      handler(parsed.success ? parsed.data : { online: current.online })
      return () => {
        watchers.delete(handler)
      }
    },
  }
  return handle
}
