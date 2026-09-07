/**
 * Offline queue — store/sync/clear actions taken while disconnected.
 * In-memory only; cleared on process restart.
 */
import { OFFLINE_QUEUE, registerHandle } from '../common/index.js'

export function registerOfflineHandlers(): void {
  registerHandle('offline:queue', (_event: unknown, args: unknown) => {
    const { action, payload } = args as { action: string; payload: unknown }

    const id = `offline-${Date.now()}`
    OFFLINE_QUEUE.set(id, {
      id,
      action,
      payload,
      timestamp: Date.now(),
      synced: false,
    })

    return { ok: true, id, queued: OFFLINE_QUEUE.size }
  })

  registerHandle('offline:get-queue', () => {
    return [...OFFLINE_QUEUE.values()].map(q => ({
      id: q.id,
      action: q.action,
      timestamp: q.timestamp,
      synced: q.synced,
    }))
  })

  registerHandle('offline:sync', async (_event: unknown) => {
    const pending = [...OFFLINE_QUEUE.values()].filter(q => !q.synced)
    const synced: string[] = []

    for (const item of pending) {
      item.synced = true
      synced.push(item.id)
    }

    return { ok: true, synced: synced.length, ids: synced }
  })

  registerHandle('offline:clear', (_event: unknown) => {
    const count = OFFLINE_QUEUE.size
    OFFLINE_QUEUE.clear()
    return { ok: true, cleared: count }
  })
}
