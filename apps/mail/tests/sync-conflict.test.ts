import { describe, expect, it } from 'vitest'
import type { MailEvent } from '@genoffice/mail-protocols'
import { SyncEngine } from '../src/main/sync/engine'
import { FakeSyncStore, ScriptedDriver, testEnvelope } from './sync-fixtures'

function harness(intentOutcomes: Array<'acked' | 'stale-modseq' | 'failed'> = []): {
  store: FakeSyncStore
  driver: ScriptedDriver
  engine: SyncEngine
  events: MailEvent[]
} {
  const store = new FakeSyncStore()
  const driver = new ScriptedDriver({ accountId: 'a', kind: 'imap', intentOutcomes })
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

describe('flag-move conflict, server wins (spec case: flag-move-conflict-server-wins)', () => {
  it('fails the intent with a conflict event and keeps server state', async () => {
    const { store, driver, engine, events } = harness(['stale-modseq', 'stale-modseq'])
    // Server-side change lands first: the message is flagged server-side.
    store.upsertEnvelopes([testEnvelope('u-1', { flags: { flagged: true } })])
    // Offline intent: clear the flag with a stale modseq expectation.
    const intentId = store.queueIntent({
      accountId: 'a',
      kind: 'flag',
      targetUid: 'u-1',
      expectedModseq: '5',
      payloadJson: '{"flagged":false,"folderId":"inbox"}',
    })

    await engine.replayQueuedIntents('a')

    expect(store.intentState(intentId)).toBe('failed')
    // One refetch between the two attempts, then the conflict surfaces.
    expect(driver.refetched).toEqual(['u-1'])
    const conflict = events.find((event) => event.kind === 'conflict')
    expect(conflict).toMatchObject({ accountId: 'a', folderId: 'inbox' })
    // Server state stays visible: the flag the server set is still set.
    expect(store.getThread('thread-u-1')[0].flags.flagged).toBe(true)
  })

  it('acks a clean intent with no conflict event', async () => {
    const { store, engine, events } = harness(['acked'])
    const intentId = store.queueIntent({
      accountId: 'a',
      kind: 'move',
      targetUid: 'u-2',
      payloadJson: '{"toFolderId":"archive","folderId":"inbox"}',
    })
    await engine.replayQueuedIntents('a')
    expect(store.intentState(intentId)).toBe('acked')
    expect(events.some((event) => event.kind === 'conflict')).toBe(false)
  })

  it('recovers after one stale modseq with a refetch and retry', async () => {
    const { store, driver, engine, events } = harness(['stale-modseq', 'acked'])
    const intentId = store.queueIntent({
      accountId: 'a',
      kind: 'flag',
      targetUid: 'u-3',
      expectedModseq: '5',
      payloadJson: '{"seen":true,"folderId":"inbox"}',
    })
    await engine.replayQueuedIntents('a')
    expect(store.intentState(intentId)).toBe('acked')
    expect(driver.refetched).toEqual(['u-3'])
    expect(events.some((event) => event.kind === 'conflict')).toBe(false)
  })

  it('fails transport errors with a conflict event, never a silent merge', async () => {
    const { store, engine, events } = harness(['failed'])
    const intentId = store.queueIntent({
      accountId: 'a',
      kind: 'delete',
      targetUid: 'u-4',
      payloadJson: '{"folderId":"inbox"}',
    })
    await engine.replayQueuedIntents('a')
    expect(store.intentState(intentId)).toBe('failed')
    expect(events.some((event) => event.kind === 'conflict')).toBe(true)
  })

  it('replays intents FIFO', async () => {
    const { store, driver, engine } = harness(['acked', 'acked'])
    const first = store.queueIntent({
      accountId: 'a',
      kind: 'flag',
      targetUid: 'u-5',
      payloadJson: '{"folderId":"inbox"}',
    })
    const second = store.queueIntent({
      accountId: 'a',
      kind: 'flag',
      targetUid: 'u-6',
      payloadJson: '{"folderId":"inbox"}',
    })
    await engine.replayQueuedIntents('a')
    expect(driver.intentAttempts.map((intent) => intent.id)).toEqual([first, second])
    expect(store.intentState(first)).toBe('acked')
    expect(store.intentState(second)).toBe('acked')
  })
})
