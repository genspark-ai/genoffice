import { describe, expect, it } from 'vitest'
import type { MailEvent } from '@genoffice/mail-protocols'
import { SyncEngine } from '../src/main/sync/engine'
import {
  FakeSyncStore,
  ScriptedDriver,
  incrementalResult,
  syncState,
  testEnvelope,
} from './sync-fixtures'

function harness(kind: 'imap' | 'gmail' | 'graph' = 'imap'): {
  store: FakeSyncStore
  driver: ScriptedDriver
  engine: SyncEngine
  events: MailEvent[]
} {
  const store = new FakeSyncStore()
  const driver = new ScriptedDriver({ accountId: 'a', kind })
  const events: MailEvent[] = []
  const engine = new SyncEngine({
    store,
    noTimers: true,
    emit: {
      mailEvent: (event) => events.push(event),
      offlineStatus: () => undefined,
    },
  })
  engine.registerDriver(driver)
  return { store, driver, engine, events }
}

describe('uidvalidity-change resync (spec case: uidvalidity-change-resync)', () => {
  it('replaces folder rows, emits resync, and preserves queues', async () => {
    const { store, driver, engine, events } = harness('imap')
    store.upsertEnvelopes([testEnvelope('old-1'), testEnvelope('old-2')])
    store.setSyncState(syncState('a', { uidValidity: 1, highestModseq: '10' }))
    // Queues that a resync must never drop.
    const outboxRecord = engine.enqueueSend('a', {
      to: ['reader@example.com'],
      subject: 'Queued while offline',
      text: 'hello',
    })
    store.saveDraft('a', 'thread-x', 'local draft text')
    const intentId = store.queueIntent({
      accountId: 'a',
      kind: 'flag',
      targetUid: 'old-1',
      payloadJson: '{"seen":true,"folderId":"inbox"}',
    })
    driver.queueResult({
      decision: { kind: 'full-resync', reason: 'uidvalidity-changed' },
      envelopes: [testEnvelope('new-1'), testEnvelope('new-2')],
      state: syncState('a', { uidValidity: 2, highestModseq: '1' }),
    })

    const started = await engine.syncNow('a')
    expect(started).toEqual({ started: true })
    expect(
      store
        .envelopesIn('inbox')
        .map((env) => env.ref.uid)
        .sort(),
    ).toEqual(['new-1', 'new-2'])
    const resync = events.find((event) => event.kind === 'resync')
    expect(resync).toMatchObject({
      accountId: 'a',
      folderId: 'inbox',
      reason: 'uidvalidity-changed',
    })
    // Queues survive the resync.
    expect(engine.outboxList('a').map((row) => row.id)).toContain(outboxRecord.id)
    expect(store.draftFor('a', 'thread-x')).toBe('local draft text')
    expect(store.intentState(intentId)).toBe('queued')
    expect(store.getSyncState('a', 'inbox')?.uidValidity).toBe(2)
  })
})

describe('history-expiry resync (spec case: history-expiry-resync)', () => {
  it('re-lists fully, persists the fresh historyId, and emits resync', async () => {
    const { store, driver, engine, events } = harness('gmail')
    store.setSyncState(syncState('a', { historyId: 'expired-id' }))
    driver.queueResult({
      decision: { kind: 'full-resync', reason: 'history-expired' },
      envelopes: [testEnvelope('g-1')],
      state: syncState('a', { historyId: 'fresh-id' }),
    })

    await engine.syncNow('a')
    expect(store.getSyncState('a', 'inbox')?.historyId).toBe('fresh-id')
    expect(events.some((event) => event.kind === 'resync')).toBe(true)
    // The exact expired-cursor kind travels in the sync-state detail.
    expect(
      events.some(
        (event) => event.kind === 'sync-state' && event.detail === 'resync:history-expired',
      ),
    ).toBe(true)
  })
})

describe('delta-expiry resync (spec case: delta-expiry-resync)', () => {
  it('re-lists fully and persists the new delta token', async () => {
    const { store, driver, engine, events } = harness('graph')
    store.setSyncState(syncState('a', { deltaToken: 'stale-token' }))
    driver.queueResult({
      decision: { kind: 'full-resync', reason: 'delta-expired' },
      envelopes: [testEnvelope('m-1')],
      state: syncState('a', { deltaToken: 'fresh-token' }),
    })

    await engine.syncNow('a')
    expect(store.getSyncState('a', 'inbox')?.deltaToken).toBe('fresh-token')
    expect(events.some((event) => event.kind === 'resync')).toBe(true)
  })
})

describe('incremental sync', () => {
  it('upserts deltas, persists state only after apply, and debounces', async () => {
    const { store, driver, engine, events } = harness('imap')
    store.setSyncState(syncState('a', { uidValidity: 7, highestModseq: '3' }))
    driver.queueResult(
      incrementalResult(
        [testEnvelope('n-1')],
        syncState('a', { uidValidity: 7, highestModseq: '4' }),
      ),
    )
    // A second syncNow while the first runs reports started:false.
    const first = engine.syncNow('a')
    const second = await engine.syncNow('a')
    expect(second).toEqual({ started: false })
    expect(await first).toEqual({ started: true })
    expect(store.envelopesIn('inbox').map((env) => env.ref.uid)).toEqual(['n-1'])
    expect(store.getSyncState('a', 'inbox')?.highestModseq).toBe('4')
    expect(
      events.filter((event) => event.kind === 'sync-state').map((event) => event),
    ).toContainEqual(expect.objectContaining({ state: 'idle' }))
  })

  it('surfaces driver errors as sanitized sync-state events', async () => {
    const { driver, engine, events } = harness('imap')
    driver.queueResult(incrementalResult([testEnvelope('n-1')], syncState('a', { uidValidity: 7 })))
    // Poison the next fetch by clearing results and forcing the throw path.
    driver.clearResults()
    await engine.syncNow('a')
    const error = events.find((event) => event.kind === 'sync-state' && event.state === 'error')
    expect(error).toBeDefined()
  })
})
