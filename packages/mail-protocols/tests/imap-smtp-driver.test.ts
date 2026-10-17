/**
 * Lane B fixtures: round-trip matrix row (IMAP+SMTP), UIDVALIDITY resync,
 * offline intents, send caps, IDLE/poll watch paths. All against the local
 * fake IMAP/SMTP staging servers; no network, no real credentials.
 */
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { unlinkSync, writeFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import {
  ImapSmtpDriver,
  ImapSmtpError,
  InMemoryIntentStore,
  InMemoryOutboxStore,
  InMemorySyncStateStore,
  MAX_ATTACHMENT_BYTES,
  MAX_RECIPIENTS,
  MAX_TOTAL_BYTES,
  type ImapSmtpSendCaps,
} from '../src/imap-smtp-driver.js'
import { OUTGOING_WARN_BYTES } from '../src/attachments.js'
import { FakeImapServer, FakeSmtpServer } from '../src/imap-fakes.js'
import type { MailEvent } from '../src/types.js'

const sleep = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms))

async function waitFor(check: () => boolean, timeoutMs = 3000): Promise<void> {
  const start = Date.now()
  for (;;) {
    if (check()) return
    if (Date.now() - start > timeoutMs) throw new Error('waitFor timed out')
    await sleep(10)
  }
}

function makeHarness(options?: {
  idle?: boolean
  condstore?: boolean
  caps?: ImapSmtpSendCaps
  pollIntervalMs?: number
}): {
  imap: FakeImapServer
  smtp: FakeSmtpServer
  syncStore: InMemorySyncStateStore
  intentStore: InMemoryIntentStore
  outbox: InMemoryOutboxStore
  events: MailEvent[]
  driver: ImapSmtpDriver
} {
  const imap = new FakeImapServer({ idle: options?.idle, condstore: options?.condstore })
  const smtp = new FakeSmtpServer()
  const syncStore = new InMemorySyncStateStore()
  const intentStore = new InMemoryIntentStore()
  const outbox = new InMemoryOutboxStore()
  const events: MailEvent[] = []
  const driver = new ImapSmtpDriver(
    {
      accountId: 'test-account',
      imap: {
        host: 'imap.example.local',
        port: 993,
        secure: true,
        auth: { user: 'test-user', password: 'test-pass' },
      },
      smtp: {
        host: 'smtp.example.local',
        port: 465,
        secure: true,
        auth: { user: 'test-user', password: 'test-pass' },
        from: 'test-user@example.local',
      },
      pollIntervalMs: options?.pollIntervalMs,
      sendCaps: options?.caps,
    },
    {
      createImapClient: () => imap,
      createSmtpTransport: () => smtp,
      syncStore,
      intentStore,
      outbox,
      onEvent: (ev) => events.push(ev),
    },
  )
  return { imap, smtp, syncStore, intentStore, outbox, events, driver }
}

describe('round-trip matrix row (imap+smtp)', () => {
  it('sends via SMTP, delivers to IMAP, and reads back the same envelope', async () => {
    const { imap, smtp, driver, outbox } = makeHarness()
    await driver.connect()
    const result = await driver.send({
      to: ['friend@example.local'],
      subject: 'Round trip',
      text: 'Hello over the fake wire',
    })
    expect(result.outboxId).toMatch(/^outbox-/)
    expect(result.messageId).toMatch(/fake-/)
    expect(result.warnings).toEqual([])
    expect(smtp.sent).toHaveLength(1)
    expect(outbox.entries.get(result.outboxId)?.state).toBe('sent')

    smtp.deliverTo(imap, 'INBOX')
    const envelopes = await driver.list('INBOX', { limit: 50 })
    expect(envelopes).toHaveLength(1)
    const env = envelopes[0]
    expect(env?.subject).toBe('Round trip')
    expect(env?.from).toEqual(['test-user@example.local'])
    expect(env?.to).toEqual(['friend@example.local'])
    expect(env?.hasAttach).toBe(false)
    expect(env?.ref.threadId).toBeTruthy()
    if (!env) throw new Error('expected one envelope')
    const body = await driver.fetch(env.ref)
    expect(body.text).toBe('Hello over the fake wire')
    await driver.disconnect()
  })
})

describe('send caps', () => {
  it('defaults match the spec: 25 MiB warning, hard caps above it', () => {
    expect(OUTGOING_WARN_BYTES).toBe(25 * 1024 * 1024)
    expect(MAX_TOTAL_BYTES).toBeGreaterThan(OUTGOING_WARN_BYTES)
    expect(MAX_ATTACHMENT_BYTES).toBe(25 * 1024 * 1024)
    expect(MAX_RECIPIENTS).toBe(100)
  })

  it('warns above the warning threshold but still sends', async () => {
    const { smtp, driver } = makeHarness({ caps: { warnTotalBytes: 10, maxTotalBytes: 10_000 } })
    await driver.connect()
    const result = await driver.send({
      to: ['friend@example.local'],
      subject: 'Big-ish',
      text: 'x'.repeat(64),
    })
    expect(result.warnings).toHaveLength(1)
    expect(smtp.sent).toHaveLength(1)
    await driver.disconnect()
  })

  it('rejects too many recipients', async () => {
    const { smtp, driver } = makeHarness()
    await driver.connect()
    const to = Array.from({ length: MAX_RECIPIENTS + 1 }, (_, i) => `u${i}@example.local`)
    await expect(driver.send({ to, subject: 'Spam', text: 'nope' })).rejects.toMatchObject({
      code: 'too-many-recipients',
    })
    expect(smtp.sent).toHaveLength(0)
    await driver.disconnect()
  })

  it('rejects an oversized attachment and an oversized total', async () => {
    const { driver } = makeHarness({ caps: { maxAttachmentBytes: 8, maxTotalBytes: 64 } })
    await driver.connect()
    await expect(
      driver.send({
        to: ['friend@example.local'],
        subject: 'Big file',
        text: 'see attached',
        attachments: [
          { filename: 'big.bin', mime: 'application/octet-stream', size: 16, path: '' },
        ],
      }),
    ).rejects.toMatchObject({ code: 'attachment-too-large' })
    await expect(
      driver.send({
        to: ['friend@example.local'],
        subject: 'Big body',
        text: 'x'.repeat(128),
      }),
    ).rejects.toMatchObject({ code: 'oversize' })
    await driver.disconnect()
  })

  it('streams attachments from disk paths instead of buffering them', async () => {
    const { smtp, driver } = makeHarness()
    const staged = join(tmpdir(), `mail-lane-b-${Date.now()}.bin`)
    writeFileSync(staged, Buffer.alloc(1024, 7))
    try {
      await driver.connect()
      await driver.send({
        to: ['friend@example.local'],
        subject: 'Staged file',
        text: 'streamed',
        attachments: [
          {
            filename: 'staged.bin',
            mime: 'application/octet-stream',
            size: 1024,
            path: staged,
          },
        ],
      })
      const captured = smtp.sent[0]?.attachments[0]
      expect(captured?.path).toBe(staged)
      expect(captured?.content).toBeUndefined()
    } finally {
      unlinkSync(staged)
      await driver.disconnect()
    }
  })

  it('queues offline sends without touching SMTP (client wins for outbox)', async () => {
    const { smtp, driver, outbox } = makeHarness()
    const result = await driver.send({
      to: ['friend@example.local'],
      subject: 'Offline note',
      text: 'send me later',
    })
    expect(result.messageId).toBeUndefined()
    expect(smtp.sent).toHaveLength(0)
    expect(outbox.entries.get(result.outboxId)?.state).toBe('queued')
  })
})

describe('UIDVALIDITY resync', () => {
  it('re-fetches the folder and never drops queued outbox or intents', async () => {
    const { imap, driver, events, intentStore, outbox } = makeHarness()
    imap.seedMessage('INBOX', {
      from: ['a@example.local'],
      to: ['me@example.local'],
      subject: 'One',
    })
    imap.seedMessage('INBOX', {
      from: ['b@example.local'],
      to: ['me@example.local'],
      subject: 'Two',
    })
    await driver.connect()
    expect(await driver.list('INBOX', { limit: 50 })).toHaveLength(2)

    // Queue offline work, then go back online with a changed UIDVALIDITY.
    await driver.disconnect()
    const queuedOutbox = outbox.enqueueSend('test-account', {
      to: ['friend@example.local'],
      subject: 'Waiting',
      text: 'still here',
    })
    await driver.move({ uid: '1', folderId: 'INBOX', threadId: 'INBOX:1' }, 'ARCHIVE')
    await driver.connect()

    imap.setUidValidity('INBOX', 2)
    imap.seedMessage('INBOX', {
      from: ['c@example.local'],
      to: ['me@example.local'],
      subject: 'Three',
    })
    const after = await driver.list('INBOX', { limit: 50 })
    expect(after.length).toBeGreaterThanOrEqual(3)
    expect(events).toContainEqual({
      kind: 'resync',
      accountId: 'test-account',
      folderId: 'INBOX',
      reason: 'uidvalidity-changed',
    })
    expect(outbox.entries.get(queuedOutbox)?.state).toBe('queued')
    expect(intentStore.pendingIntents('test-account')).toHaveLength(1)
    await driver.disconnect()
  })
})

describe('CONDSTORE vs UID-range fetch', () => {
  it('uses MODSEQ deltas when the server advertises CONDSTORE', async () => {
    const { imap, driver } = makeHarness({ condstore: true })
    imap.seedMessage('INBOX', {
      from: ['a@example.local'],
      to: ['me@example.local'],
      subject: 'One',
    })
    await driver.connect()
    await driver.list('INBOX', { limit: 50 })
    const baseline = imap.fetchSinceModseqCalls
    imap.touchMessage('INBOX', '1', { seen: true })
    const delta = await driver.list('INBOX', { limit: 50 })
    expect(imap.fetchSinceModseqCalls).toBeGreaterThan(baseline)
    expect(delta).toHaveLength(1)
    expect(delta[0]?.flags.seen).toBe(true)
    await driver.disconnect()
  })

  it('falls back to UID-range fetch without CONDSTORE', async () => {
    const { imap, driver } = makeHarness({ condstore: false })
    imap.seedMessage('INBOX', {
      from: ['a@example.local'],
      to: ['me@example.local'],
      subject: 'One',
    })
    imap.seedMessage('INBOX', {
      from: ['b@example.local'],
      to: ['me@example.local'],
      subject: 'Two',
    })
    await driver.connect()
    await driver.list('INBOX', { limit: 50 })
    await driver.list('INBOX', { limit: 50 })
    expect(imap.fetchSinceModseqCalls).toBe(0)
    expect(imap.fetchRangeCalls).toBeGreaterThan(0)
    const page = await driver.list('INBOX', { afterUid: '1', limit: 50 })
    expect(page.map((e) => e.ref.uid)).toEqual(['2'])
    await driver.disconnect()
  })
})

describe('watch: IDLE vs poll', () => {
  it('pushes IDLE events where advertised', async () => {
    const { imap, driver } = makeHarness({ idle: true })
    expect(imap.supports('IDLE')).toBe(true)
    await driver.connect()
    const seen: MailEvent[] = []
    const sub = driver.watch((ev) => seen.push(ev))
    imap.simulateIncoming('INBOX', {
      from: ['a@example.local'],
      to: ['me@example.local'],
      subject: 'Ping',
    })
    expect(seen).toContainEqual({
      kind: 'messages-added',
      accountId: 'test-account',
      folderId: 'INBOX',
      refs: [{ uid: '1', folderId: 'INBOX', threadId: 'INBOX:1' }],
    })
    sub.stop()
    imap.simulateIncoming('INBOX', {
      from: ['b@example.local'],
      to: ['me@example.local'],
      subject: 'After stop',
    })
    expect(seen).toHaveLength(1)
    await driver.disconnect()
  })

  it('polls with backoff where IDLE is missing, then recovers', async () => {
    const { imap, driver } = makeHarness({ idle: false, pollIntervalMs: 20 })
    expect(imap.supports('IDLE')).toBe(false)
    imap.seedMessage('INBOX', {
      from: ['a@example.local'],
      to: ['me@example.local'],
      subject: 'Seed',
    })
    await driver.connect()
    const seen: MailEvent[] = []
    const sub = driver.watch((ev) => seen.push(ev))
    try {
      imap.failOnSelect = new Error('network down')
      await waitFor(() => seen.some((ev) => ev.kind === 'sync-state' && ev.state === 'error'))
      imap.failOnSelect = undefined
      const uid = imap.simulateIncoming('INBOX', {
        from: ['b@example.local'],
        to: ['me@example.local'],
        subject: 'After outage',
      })
      await waitFor(() =>
        seen.some(
          (ev) =>
            ev.kind === 'messages-added' &&
            ev.folderId === 'INBOX' &&
            ev.refs.some((r) => r.uid === uid),
        ),
      )
    } finally {
      sub.stop()
      await driver.disconnect()
    }
  })
})

describe('offline intents', () => {
  it('queues moves/flags offline and replays them idempotently', async () => {
    const { imap, driver, intentStore } = makeHarness()
    imap.seedMessage('INBOX', {
      from: ['a@example.local'],
      to: ['me@example.local'],
      subject: 'One',
    })
    await driver.connect()
    await driver.list('INBOX', { limit: 50 })
    await driver.disconnect()

    await driver.setFlags({ uid: '1', folderId: 'INBOX', threadId: 'INBOX:1' }, { seen: true })
    await driver.move({ uid: '1', folderId: 'INBOX', threadId: 'INBOX:1' }, 'ARCHIVE')
    const pending = intentStore.pendingIntents('test-account')
    expect(pending).toHaveLength(2)
    for (const intent of pending) {
      expect(intent.targetUid).toBe('1')
      expect(intent.expectedModseq).toBeDefined()
    }

    await driver.connect()
    const first = await driver.drainIntents()
    expect(first.failed).toEqual([])
    expect(first.acked).toHaveLength(2)
    expect(imap.messageCount('INBOX')).toBe(0)
    expect(imap.messageCount('ARCHIVE')).toBe(1)
    const archived = await driver.list('ARCHIVE', { limit: 50 })
    expect(archived[0]?.flags.seen).toBe(true)

    const second = await driver.drainIntents()
    expect(second).toEqual({ acked: [], failed: [] })
    await driver.disconnect()
  })

  it('re-fetches and retries once on a stale modseq', async () => {
    const { imap, driver } = makeHarness()
    imap.seedMessage('INBOX', {
      from: ['a@example.local'],
      to: ['me@example.local'],
      subject: 'One',
    })
    await driver.connect()
    await driver.list('INBOX', { limit: 50 })
    await driver.disconnect()

    await driver.setFlags({ uid: '1', folderId: 'INBOX', threadId: 'INBOX:1' }, { flagged: true })
    imap.touchMessage('INBOX', '1', { seen: true })
    imap.staleOnce.add('1')
    const rangeCalls = imap.fetchRangeCalls

    await driver.connect()
    const result = await driver.drainIntents()
    expect(result.failed).toEqual([])
    expect(result.acked).toHaveLength(1)
    // Exactly one re-fetch powered the single retry.
    expect(imap.fetchRangeCalls).toBeGreaterThan(rangeCalls)
    const envelopes = await driver.list('INBOX', { limit: 50 })
    expect(envelopes[0]?.flags.flagged).toBe(true)
    await driver.disconnect()
  })

  it('surfaces a conflict when the modseq stays stale after re-fetch', async () => {
    const { imap, driver, intentStore, events } = makeHarness()
    imap.seedMessage('INBOX', {
      from: ['a@example.local'],
      to: ['me@example.local'],
      subject: 'One',
    })
    await driver.connect()
    await driver.list('INBOX', { limit: 50 })
    await driver.disconnect()

    await driver.setFlags({ uid: '1', folderId: 'INBOX', threadId: 'INBOX:1' }, { flagged: true })
    expect(intentStore.pendingIntents('test-account')).toHaveLength(1)
    imap.alwaysStale = true

    await driver.connect()
    const result = await driver.drainIntents()
    expect(result.acked).toEqual([])
    expect(result.failed).toHaveLength(1)
    expect(events).toContainEqual({
      kind: 'conflict',
      accountId: 'test-account',
      folderId: 'INBOX',
      uid: '1',
      detail: expect.any(String) as string,
    })
    expect(intentStore.pendingIntents('test-account')).toHaveLength(0)
    imap.alwaysStale = false
    await driver.disconnect()
  })
})

describe('server search expansion', () => {
  it('delegates explicit server SEARCH to IMAP', async () => {
    const { imap, driver } = makeHarness()
    imap.seedMessage('INBOX', {
      from: ['a@example.local'],
      to: ['me@example.local'],
      subject: 'Invoice',
    })
    imap.seedMessage('INBOX', {
      from: ['b@example.local'],
      to: ['me@example.local'],
      subject: 'Lunch',
    })
    await driver.connect()
    const hits = await driver.search('INBOX', { text: 'invoice' })
    expect(hits.map((e) => e.subject)).toEqual(['Invoice'])
    await driver.disconnect()
  })

  it('honors the extended to/since/before fields server-side', async () => {
    const { imap, driver } = makeHarness()
    imap.seedMessage('INBOX', {
      from: ['a@example.local'],
      to: ['me@example.local'],
      subject: 'Old note',
      date: '2026-01-05T10:00:00.000Z',
    })
    imap.seedMessage('INBOX', {
      from: ['b@example.local'],
      to: ['boss@example.local'],
      subject: 'Q1 report',
      date: '2026-02-10T10:00:00.000Z',
    })
    imap.seedMessage('INBOX', {
      from: ['c@example.local'],
      to: ['boss@example.local'],
      subject: 'Q3 report',
      date: '2026-05-10T10:00:00.000Z',
    })
    await driver.connect()
    const hits = await driver.search('INBOX', {
      to: 'boss@example.local',
      since: '2026-02-01',
      before: '2026-04-01',
    })
    expect(hits.map((e) => e.subject)).toEqual(['Q1 report'])
    await driver.disconnect()
  })
})

describe('error shape', () => {
  it('throws typed offline errors when disconnected', async () => {
    const { driver } = makeHarness()
    const err = await driver.list('INBOX', { limit: 10 }).catch((e: unknown) => e)
    expect(err).toBeInstanceOf(ImapSmtpError)
    expect((err as ImapSmtpError).code).toBe('offline')
  })
})
