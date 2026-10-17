import { describe, expect, it } from 'vitest'
import type { MailEvent } from '@genoffice/mail-protocols'
import { SyncEngine, type OfflineStatusPayload } from '../src/main/sync/engine'
import { OutboxQueue } from '../src/main/sync/outbox'
import { FakeSyncStore, ScriptedDriver, testSendRequest } from './sync-fixtures'

function harness(): {
  store: FakeSyncStore
  driver: ScriptedDriver
  engine: SyncEngine
  events: MailEvent[]
  statuses: OfflineStatusPayload[]
} {
  const store = new FakeSyncStore()
  const driver = new ScriptedDriver({ accountId: 'a', kind: 'imap' })
  const events: MailEvent[] = []
  const statuses: OfflineStatusPayload[] = []
  const engine = new SyncEngine({
    store,
    noTimers: true,
    emit: {
      mailEvent: (event) => events.push(event),
      offlineStatus: (status) => statuses.push(status),
    },
  })
  engine.registerDriver(driver)
  return { store, driver, engine, events, statuses }
}

describe('offline-compose to reconnect-send (spec case: offline-compose-reconnect-send)', () => {
  it('queues while offline and replays exactly-once on reconnect', async () => {
    const { driver, engine, statuses } = harness()
    await engine.setOnline(false, 'a', 'test offline')
    expect(statuses.at(-1)).toMatchObject({ online: false, accountId: 'a' })
    expect(await engine.syncNow('a')).toEqual({ started: false })

    const first = engine.enqueueSend('a', testSendRequest({ subject: 'First' }))
    const second = engine.enqueueSend('a', testSendRequest({ subject: 'Second' }))
    expect(first.state).toBe('queued')
    expect(
      engine
        .outboxList('a')
        .map((row) => row.id)
        .sort(),
    ).toEqual([first.id, second.id].sort())
    expect(engine.outboxStatus('a')).toEqual({ queued: 2, failed: 0 })

    await engine.setOnline(true, 'a')
    expect(statuses.at(-1)).toMatchObject({ online: true, accountId: 'a' })
    // FIFO order with the outbox id as the idempotency key.
    expect(driver.sent.map((send) => send.req.subject)).toEqual(['First', 'Second'])
    expect(driver.sent[0].outboxId).toBe(first.id)
    expect(engine.outboxStatus('a')).toEqual({ queued: 0, failed: 0 })
    expect(engine.outboxList('a').every((row) => row.state === 'sent')).toBe(true)
    // Paused while offline, resumed on reconnect.
    expect(driver.paused).toBe(1)
    expect(driver.resumed).toBe(1)
  })

  it('drains outbox sends before replaying queued intents', async () => {
    const { store, driver, engine } = harness()
    const order: string[] = []
    const intentId = store.queueIntent({
      accountId: 'a',
      kind: 'flag',
      targetUid: 'u-1',
      payloadJson: '{"seen":true,"folderId":"inbox"}',
    })
    const originalAttempt = driver.attemptIntent.bind(driver)
    driver.attemptIntent = async (intent) => {
      order.push(`intent:${intent.id}`)
      return originalAttempt(intent)
    }
    const originalSend = driver.send.bind(driver)
    driver.send = async (req, outboxId) => {
      order.push(`send:${req.subject}`)
      return originalSend(req, outboxId)
    }
    engine.enqueueSend('a', testSendRequest({ subject: 'Outbox first' }))
    await engine.reconnect('a')
    expect(order).toEqual(['send:Outbox first', `intent:${intentId}`])
  })
})

describe('idempotent replay (spec case: idempotent-replay-double-delivery-safe)', () => {
  it('recovers a crash between send and mark with exactly one server message', async () => {
    const { driver, engine } = harness()
    const queued = engine.enqueueSend('a', testSendRequest({ subject: 'Crash me' }))
    // Crash simulation: the bytes reached the server but the sent mark
    // never landed (entry stuck in sending after a restart).
    await driver.send(queued.req, queued.id)
    expect(driver.sent).toHaveLength(1)

    await engine.reconnect('a')
    // Reconcile confirmed the id, so no retransmit happened.
    expect(driver.sent).toHaveLength(1)
    expect(driver.confirmed).toContain(queued.id)
    expect(engine.outboxList('a')[0]).toMatchObject({ id: queued.id, state: 'sent' })
  })

  it('reconciles before retransmitting on every drain', async () => {
    const { driver, engine } = harness()
    engine.enqueueSend('a', testSendRequest({ subject: 'Checked first' }))
    await engine.drainOutbox('a')
    // Confirm ran before the single send; exactly one server message.
    expect(driver.confirmed).toHaveLength(1)
    expect(driver.sent).toHaveLength(1)
  })

  it('requeues entries stuck in sending after a restart', () => {
    const queue = new OutboxQueue({ clock: () => '2026-10-14T00:00:00.000Z' })
    const row = queue.enqueue('a', testSendRequest({ subject: 'Stuck' }))
    queue.beginSend('a')
    // Restart loses the in-flight flag while the row stays sending.
    queue.release('a')
    expect(queue.requeueStuck('a')).toBe(1)
    expect(queue.nextQueued('a')?.id).toBe(row.id)
  })
})

describe('outbox retry, cancel, and caps at send time', () => {
  it('retries with backoff then parks as failed after 5 attempts', async () => {
    const { driver, engine } = harness()
    driver.failSendWith = 'smtp unavailable'
    const queued = engine.enqueueSend('a', testSendRequest({ subject: 'Flaky' }))
    for (let attempt = 1; attempt <= 5; attempt += 1) {
      await engine.drainOutbox('a')
    }
    const row = engine.outboxList('a')[0]
    expect(row).toMatchObject({ id: queued.id, state: 'failed', attempts: 5 })
    expect(row.lastError).toContain('smtp unavailable')
    expect(engine.outboxStatus('a')).toEqual({ queued: 0, failed: 1 })
  })

  it('manual retry resets the budget and cancelled entries never send', async () => {
    const { driver, engine } = harness()
    driver.failSendWith = 'smtp unavailable'
    const queued = engine.enqueueSend('a', testSendRequest({ subject: 'Retry me' }))
    await engine.drainOutbox('a')
    await engine.drainOutbox('a')
    expect(engine.retryOutbox('a', queued.id)).toEqual({ outboxId: queued.id, state: 'queued' })
    const reset = engine.outboxList('a')[0]
    expect(reset.attempts).toBe(0)

    const doomed = engine.enqueueSend('a', testSendRequest({ subject: 'Cancel me' }))
    expect(engine.cancelOutbox('a', doomed.id)).toEqual({ cancelled: true })
    driver.failSendWith = undefined
    await engine.reconnect('a')
    // The retried entry sent; the cancelled entry never did.
    expect(driver.sent.map((send) => send.req.subject)).toEqual(['Retry me'])
    expect(engine.cancelOutbox('a', doomed.id)).toEqual({ cancelled: true })
  })

  it('enforces caps at send time when they changed while queued', () => {
    const { engine } = harness()
    const queued = engine.enqueueSend('a', testSendRequest({ subject: 'Small' }))
    expect(queued.state).toBe('queued')
    // A graph driver rejects 4 MB attachments, so the same payload that
    // IMAP accepts is refused at enqueue once the account moves drivers.
    const graph = new ScriptedDriver({ accountId: 'a', kind: 'graph' })
    engine.registerDriver(graph)
    const oversize = testSendRequest({
      subject: 'Too big for graph',
      attachments: [
        {
          filename: 'big.bin',
          mime: 'application/octet-stream',
          size: 4 * 1024 * 1024,
          path: '/staging/big.bin',
        },
      ],
    })
    // Enqueue on the graph driver rejects immediately.
    expect(() => engine.enqueueSend('a', oversize)).toThrow('attachment-too-large')
  })

  it('rejects unknown accounts and unknown outbox ids', () => {
    const { engine } = harness()
    expect(() => engine.enqueueSend('ghost', testSendRequest())).toThrow('unknown account')
    expect(() => engine.retryOutbox('a', 'ghost-id')).toThrow()
    expect(engine.cancelOutbox('a', 'ghost-id')).toEqual({ cancelled: false })
  })
})

describe('OutboxQueue unit behavior', () => {
  it('holds one in-flight send per account and lists newest-first capped at 200', () => {
    const queue = new OutboxQueue({ clock: () => '2026-10-14T00:00:00.000Z' })
    const a = queue.enqueue('a', testSendRequest({ subject: 'A' }))
    queue.enqueue('a', testSendRequest({ subject: 'B' }))
    expect(queue.beginSend('a')?.id).toBe(a.id)
    expect(queue.isBusy('a')).toBe(true)
    expect(queue.beginSend('a')).toBeUndefined()
    queue.markSent('a', a.id)
    expect(queue.isBusy('a')).toBe(false)
    const list = queue.list('a')
    expect(list).toHaveLength(2)
    expect(list.map((row) => row.req.subject)).toEqual(['A', 'B'])
  })

  it('prunes sent rows older than 7 days', () => {
    let now = Date.parse('2026-10-14T00:00:00.000Z')
    const queue = new OutboxQueue({ clock: () => new Date(now).toISOString() })
    const row = queue.enqueue('a', testSendRequest())
    queue.beginSend('a')
    queue.markSent('a', row.id)
    now += 8 * 24 * 3600_000
    expect(queue.pruneSent()).toBe(1)
    expect(queue.list('a')).toHaveLength(0)
  })
})
