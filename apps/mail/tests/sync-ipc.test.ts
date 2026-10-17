import { describe, expect, it } from 'vitest'
import {
  mailOfflineStatusSchema,
  mailOutboxCancelSchema,
  mailOutboxListSchema,
  mailOutboxRetrySchema,
  mailSyncNowSchema,
  outboxRecordSchema,
} from '../src/shared/ipc-schemas'
import { MAIL_CHANNELS } from '../src/shared/ipc'
import { SyncEngine, type OfflineStatusPayload } from '../src/main/sync/engine'
import { registerSyncHandlers, type SyncIpcBus } from '../src/main/sync/ipc'
import {
  FakeSyncStore,
  ScriptedDriver,
  incrementalResult,
  syncState,
  testEnvelope,
  testSendRequest,
} from './sync-fixtures'

function fakeBus(): SyncIpcBus & {
  handlers: Map<string, (payload: unknown) => Promise<unknown>>
  call(channel: string, payload: unknown): Promise<unknown>
} {
  const handlers = new Map<string, (payload: unknown) => Promise<unknown>>()
  return {
    handlers,
    handle: (channel, listener) => {
      handlers.set(channel, listener)
    },
    call: (channel, payload) => {
      const listener = handlers.get(channel)
      if (!listener) throw new Error(`no handler for ${channel}`)
      return listener(payload)
    },
  }
}

function wired(): {
  bus: ReturnType<typeof fakeBus>
  engine: SyncEngine
  driver: ScriptedDriver
  sent: OfflineStatusPayload[]
} {
  const store = new FakeSyncStore()
  const driver = new ScriptedDriver({ accountId: 'a', kind: 'imap' })
  const sent: OfflineStatusPayload[] = []
  const engine = new SyncEngine({ store, noTimers: true })
  engine.registerDriver(driver)
  const bus = fakeBus()
  registerSyncHandlers(bus, engine, { sendOfflineStatus: (status) => sent.push(status) })
  return { bus, engine, driver, sent }
}

describe('sync/outbox IPC channels (spec section 6)', () => {
  it('locks the five canonical channel names', () => {
    expect(MAIL_CHANNELS.syncNow).toBe('mail:sync-now')
    expect(MAIL_CHANNELS.outboxList).toBe('mail:outbox-list')
    expect(MAIL_CHANNELS.outboxRetry).toBe('mail:outbox-retry')
    expect(MAIL_CHANNELS.outboxCancel).toBe('mail:outbox-cancel')
    expect(MAIL_CHANNELS.offlineStatus).toBe('mail:offline-status')
  })

  it('validates sync/outbox payloads strictly', () => {
    expect(mailSyncNowSchema.safeParse({ accountId: 'a' }).success).toBe(true)
    expect(mailSyncNowSchema.safeParse({ accountId: 'a', folderId: 'inbox' }).success).toBe(true)
    expect(mailSyncNowSchema.safeParse({ accountId: '' }).success).toBe(false)
    expect(mailSyncNowSchema.safeParse({ accountId: 'a', bogus: 1 }).success).toBe(false)
    expect(mailOutboxListSchema.safeParse({ accountId: 'a' }).success).toBe(true)
    expect(mailOutboxRetrySchema.safeParse({ accountId: 'a', outboxId: 'o-1' }).success).toBe(true)
    expect(mailOutboxCancelSchema.safeParse({ accountId: 'a', outboxId: 'o-1' }).success).toBe(true)
    expect(mailOutboxCancelSchema.safeParse({ accountId: 'a' }).success).toBe(false)
    expect(mailOfflineStatusSchema.safeParse({ online: true }).success).toBe(true)
    expect(
      mailOfflineStatusSchema.safeParse({ online: false, accountId: 'a', detail: 'x' }).success,
    ).toBe(true)
    expect(mailOfflineStatusSchema.safeParse({ online: 'yes' }).success).toBe(false)
  })

  it('validates outbox records with the five UI states', () => {
    const base = {
      id: 'o-1',
      accountId: 'a',
      req: testSendRequest(),
      attempts: 0,
      queuedAt: '2026-10-14T00:00:00.000Z',
    }
    for (const state of ['queued', 'sending', 'failed', 'sent', 'cancelled']) {
      expect(outboxRecordSchema.safeParse({ ...base, state }).success).toBe(true)
    }
    expect(outboxRecordSchema.safeParse({ ...base, state: 'bogus' }).success).toBe(false)
  })

  it('runs sync-now through the handler with debounce', async () => {
    const { bus, driver } = wired()
    driver.queueResult(incrementalResult([testEnvelope('h-1')], syncState('a', { uidValidity: 1 })))
    expect(await bus.call(MAIL_CHANNELS.syncNow, { accountId: 'a' })).toEqual({ started: true })
    expect(driver.fetchCalls).toBe(1)
    await expect(bus.call(MAIL_CHANNELS.syncNow, { accountId: '' })).rejects.toThrow()
    await expect(bus.call(MAIL_CHANNELS.syncNow, { accountId: 'ghost' })).rejects.toThrow()
  })

  it('lists, retries, and cancels outbox entries through handlers', async () => {
    const { bus, engine } = wired()
    const queued = engine.enqueueSend('a', testSendRequest({ subject: 'Via IPC' }))
    const listed = (await bus.call(MAIL_CHANNELS.outboxList, { accountId: 'a' })) as Array<{
      id: string
    }>
    expect(listed.map((row) => row.id)).toEqual([queued.id])
    expect(
      await bus.call(MAIL_CHANNELS.outboxCancel, { accountId: 'a', outboxId: queued.id }),
    ).toEqual({ cancelled: true })
    // Retry the cancelled entry, then fail the retry shape checks.
    expect(
      await bus.call(MAIL_CHANNELS.outboxRetry, { accountId: 'a', outboxId: queued.id }),
    ).toEqual({ outboxId: queued.id, state: 'queued' })
    await expect(
      bus.call(MAIL_CHANNELS.outboxRetry, { accountId: 'a', outboxId: 'ghost' }),
    ).rejects.toThrow()
    await expect(bus.call(MAIL_CHANNELS.outboxList, { bogus: 1 })).rejects.toThrow()
  })

  it('replays offline status on subscribe and broadcasts transitions', async () => {
    const { engine, sent } = wired()
    const bus = fakeBus()
    const handle = registerSyncHandlers(bus, engine)
    const seen: OfflineStatusPayload[] = []
    const unsub = handle.onOfflineStatus((status) => seen.push(status))
    // Replay: online before any transition.
    expect(seen).toEqual([{ online: true }])
    await engine.setOnline(false, 'a', 'cable pulled')
    expect(seen.at(-1)).toMatchObject({ online: false, accountId: 'a' })
    expect(sent.at(-1)).toMatchObject({ online: false, accountId: 'a' })
    unsub()
    await engine.setOnline(true, 'a')
    expect(seen).toHaveLength(2)
  })
})
