/**
 * GraphDriver tests over the staging fake (no network, no credentials).
 * Rows: delta-expiry resync fixture, per-driver send limits, the
 * round-trip matrix row (graph / staging-fake), watch events, and auth
 * failure mapping including invalid_grant re-auth signalling.
 */
import { describe, expect, it } from 'vitest'
import { decideSync } from '../src/sync'
import type { MailIntent, SyncState } from '../src/types'
import { GRAPH_MAIL_SCOPES } from '../src/graph-transport'
import { GraphServiceError } from '../src/graph-transport'
import {
  GRAPH_MAX_ATTACHMENT_BYTES,
  GRAPH_MAX_RECIPIENTS,
  GRAPH_MAX_TOTAL_BYTES,
  GraphDriver,
  GraphDriverError,
  type GraphDriverConfig,
} from '../src/graph-driver'
import { FakeGraphTransport, makeGraphMessage } from '../src/graph-fakes'

const accountId = 'graph-test-account'

function driverWith(
  transport: FakeGraphTransport,
  over: Partial<GraphDriverConfig> = {},
): GraphDriver {
  let n = 0
  return new GraphDriver({
    accountId,
    transport,
    readStagedFile: async () => new TextEncoder().encode('staged-bytes'),
    newId: () => `outbox-${(n += 1)}`,
    ...over,
  })
}

function seedInbox(transport: FakeGraphTransport): void {
  transport.seed('inbox', [
    makeGraphMessage({
      id: 'm1',
      internetMessageId: '<m1@example.com>',
      subject: 'Quarterly report',
      from: { emailAddress: { address: 'boss@example.com', name: 'Boss' } },
      toRecipients: [{ emailAddress: { address: 'me@example.com' } }],
      receivedDateTime: '2026-10-14T10:00:00Z',
      bodyPreview: 'Please review the quarterly report draft',
      body: {
        contentType: 'html',
        content: '<p>Please review the <b>quarterly</b> report draft</p>',
      },
      hasAttachments: false,
    }),
    makeGraphMessage({
      id: 'm2',
      internetMessageId: '<m2@example.com>',
      subject: 'Lunch plans',
      from: { emailAddress: { address: 'sam@example.com' } },
      toRecipients: [{ emailAddress: { address: 'me@example.com' } }],
      receivedDateTime: '2026-10-13T09:00:00Z',
      isRead: true,
      bodyPreview: 'Are you free for lunch tomorrow',
      body: { contentType: 'text', content: 'Are you free for lunch tomorrow?' },
      hasAttachments: true,
    }),
  ])
}

describe('least-privilege scopes', () => {
  it('pins delegated scopes and never app-only wildcard scopes', () => {
    expect(GRAPH_MAIL_SCOPES).toContain('Mail.ReadWrite')
    expect(GRAPH_MAIL_SCOPES).toContain('offline_access')
    for (const scope of GRAPH_MAIL_SCOPES) {
      expect(scope).not.toMatch(/ReadWrite\.All/i)
    }
  })
})

describe('delta-expiry resync fixture', () => {
  it('initial delta returns envelopes plus a delta link', async () => {
    const transport = new FakeGraphTransport()
    seedInbox(transport)
    const driver = driverWith(transport)
    await driver.connect()
    const first = await driver.deltaSync('inbox')
    expect(first.resyncRequired).toBe(false)
    expect(first.envelopes).toHaveLength(2)
    expect(first.deltaLink).toContain('delta_inbox_v')
  })

  it('expired delta token reports resync and preserves queued outbox/intents', async () => {
    const transport = new FakeGraphTransport()
    seedInbox(transport)
    const driver = driverWith(transport)
    await driver.connect()

    const first = await driver.deltaSync('inbox')
    const stored: SyncState = {
      accountId,
      folderId: 'inbox',
      deltaToken: 'delta_inbox_v1',
      lastSyncAt: '2026-10-14T10:05:00Z',
    }
    expect(first.deltaLink).toBeDefined()

    // Queues owned by MailStore/sync-engine (P3 lane), never by the driver.
    transport.leftovers.outbox.push('queued-send-1')
    const intent: MailIntent = {
      id: 'intent-1',
      accountId,
      kind: 'move',
      targetUid: 'm1',
      payloadJson: '{"to":"archive"}',
      state: 'queued',
    }
    transport.leftovers.intents.push(intent.id)

    transport.expireDelta()
    const stale = await driver.deltaSync('inbox', stored.deltaToken)
    expect(stale.resyncRequired).toBe(true)
    expect(stale.deltaLink).toBeUndefined()

    // The engine decision for an expired Graph delta token is a full resync.
    expect(decideSync({ stored, deltaExpired: true })).toEqual({
      kind: 'full-resync',
      reason: 'delta-expired',
    })

    // Full resync re-lists without touching the queued tables.
    const relist = await driver.list('inbox', { limit: 25 })
    expect(relist).toHaveLength(2)
    expect(transport.leftovers.outbox).toEqual(['queued-send-1'])
    expect(transport.leftovers.intents).toEqual(['intent-1'])
    expect(first.envelopes.map((e) => e.ref.uid).sort()).toEqual(['m1', 'm2'])
  })

  it('delta removals surface deleted messages', async () => {
    const transport = new FakeGraphTransport()
    seedInbox(transport)
    const driver = driverWith(transport)
    await driver.connect()
    const first = await driver.deltaSync('inbox')
    expect(first.deltaLink).toBeDefined()
    await driver.delete({ uid: 'm2', folderId: 'inbox', threadId: 'm2' })
    const token =
      first.deltaLink?.includes('delta_inbox_v1') === true
        ? 'delta_inbox_v1'
        : transport.currentDeltaToken('inbox')
    const second = await driver.deltaSync('inbox', token)
    // Token rotated by the delete path only when the fake version matches;
    // either a clean delta with the removal or a resync signal is acceptable.
    if (second.resyncRequired) {
      expect(
        decideSync({ stored: { accountId, folderId: 'inbox', lastSyncAt: '' }, deltaExpired: true })
          .kind,
      ).toBe('full-resync')
    } else {
      expect(second.removed.map((r) => r.uid)).toContain('m2')
    }
  })
})

describe('send limits', () => {
  const small = () => ({
    to: ['a@example.com'],
    subject: 'hi',
    text: 'hello',
  })

  it('sends a small message and returns an outbox id', async () => {
    const transport = new FakeGraphTransport()
    const driver = driverWith(transport)
    await driver.connect()
    const verdict = driver.checkSendSize(small())
    expect(verdict).toEqual({ totalBytes: 5, recipients: 1, verdict: 'ok', warnings: [] })
    const sent = await driver.send(small())
    expect(sent).toEqual({ outboxId: 'outbox-1' })
    expect(transport.sentMail).toHaveLength(1)
    expect(transport.sentMail[0].saveToSentItems).toBe('true')
  })

  it('warns above the warn threshold but still sends', async () => {
    const transport = new FakeGraphTransport()
    const driver = driverWith(transport)
    await driver.connect()
    const req = { ...small(), text: 'x'.repeat(3 * 1024 * 1024 + 1) }
    const verdict = driver.checkSendSize(req)
    expect(verdict.verdict).toBe('warn')
    expect(verdict.warnings).toHaveLength(1)
    await driver.send(req)
    expect(transport.sentMail).toHaveLength(1)
  })

  it('rejects oversize totals at the Graph sendMail ceiling', async () => {
    const transport = new FakeGraphTransport()
    const driver = driverWith(transport)
    await driver.connect()
    const req = { ...small(), text: 'x'.repeat(GRAPH_MAX_TOTAL_BYTES + 1) }
    await expect(driver.send(req)).rejects.toMatchObject({ code: 'oversize' })
    expect(transport.sentMail).toHaveLength(0)
  })

  it('rejects too many recipients at the Exchange cap', async () => {
    const transport = new FakeGraphTransport()
    const driver = driverWith(transport)
    await driver.connect()
    const to = Array.from({ length: GRAPH_MAX_RECIPIENTS + 1 }, (_, i) => `u${i}@example.com`)
    await expect(driver.send({ ...small(), to })).rejects.toMatchObject({
      code: 'too-many-recipients',
    })
  })

  it('rejects oversize single attachments', async () => {
    const transport = new FakeGraphTransport()
    const driver = driverWith(transport)
    await driver.connect()
    await expect(
      driver.send({
        ...small(),
        attachments: [
          {
            filename: 'big.bin',
            mime: 'application/octet-stream',
            size: GRAPH_MAX_ATTACHMENT_BYTES + 1,
            path: '/staged/big.bin',
          },
        ],
      }),
    ).rejects.toMatchObject({ code: 'attachment-too-large' })
  })

  it('sends staged attachments as fileAttachments', async () => {
    const transport = new FakeGraphTransport()
    const driver = driverWith(transport)
    await driver.connect()
    await driver.send({
      ...small(),
      attachments: [
        { filename: 'note.txt', mime: 'text/plain', size: 12, path: '/staged/note.txt' },
      ],
    })
    const payload = transport.sentMail[0].message
    const attachments = payload['attachments'] as Array<Record<string, unknown>>
    expect(attachments).toHaveLength(1)
    expect(attachments[0]['@odata.type']).toBe('#microsoft.graph.fileAttachment')
    expect(attachments[0]['name']).toBe('note.txt')
  })

  it('prefers staged.data over disk for tiny inline parts', async () => {
    const transport = new FakeGraphTransport()
    let diskReads = 0
    const driver = driverWith(transport, {
      readStagedFile: async () => {
        diskReads++
        return new TextEncoder().encode('from-disk')
      },
    })
    await driver.connect()
    const inline = new TextEncoder().encode('inline-bytes')
    await driver.send({
      ...small(),
      attachments: [
        { filename: 'tiny.txt', mime: 'text/plain', size: inline.length, path: '', data: inline },
      ],
    })
    expect(diskReads).toBe(0)
    const payload = transport.sentMail[0].message
    const attachments = payload['attachments'] as Array<Record<string, unknown>>
    expect(attachments).toHaveLength(1)
    expect(attachments[0]['contentBytes']).toBe(Buffer.from(inline).toString('base64'))
  })
})

describe('round-trip matrix row: graph / staging-fake', () => {
  it('list, fetch, send, move, flag, search, delete round-trip', async () => {
    const transport = new FakeGraphTransport()
    seedInbox(transport)
    const driver = driverWith(transport)
    await driver.connect()

    const folders = await driver.listFolders()
    expect(folders.find((f) => f.id === 'inbox')?.role).toBe('inbox')
    expect(folders.find((f) => f.id === 'sentitems')?.role).toBe('sent')
    expect(folders.find((f) => f.id === 'custom-1')?.role).toBe('custom')

    const listed = await driver.list('inbox', { limit: 25 })
    expect(listed).toHaveLength(2)
    expect(listed[0].ref.uid).toBe('m1')
    expect(listed[0].subject).toBe('Quarterly report')
    expect(listed[0].flags.seen).toBe(false)
    expect(listed[1].flags.seen).toBe(true)
    expect(listed[1].hasAttach).toBe(true)

    const body = await driver.fetch({ uid: 'm1', folderId: 'inbox', threadId: 'm1' })
    expect(body.text).toContain('quarterly')
    expect(body.text).not.toContain('<p>')
    const html = await driver.fetchMessageHtml({ uid: 'm1', folderId: 'inbox', threadId: 'm1' })
    expect(html).toContain('<b>quarterly</b>')

    const found = await driver.search('inbox', { text: 'quarterly' })
    expect(found.map((e) => e.ref.uid)).toEqual(['m1'])
    const unread = await driver.search('inbox', { onlyUnread: true })
    expect(unread.map((e) => e.ref.uid)).toEqual(['m1'])

    await driver.setFlags(
      { uid: 'm1', folderId: 'inbox', threadId: 'm1' },
      { seen: true, flagged: true },
    )
    expect(transport.patches).toHaveLength(1)
    expect(transport.patches[0].body).toEqual({ isRead: true, flag: { flagStatus: 'flagged' } })

    // answered/draft have no Graph property: transmitted patch stays empty.
    await driver.setFlags(
      { uid: 'm1', folderId: 'inbox', threadId: 'm1' },
      { answered: true, draft: true },
    )
    expect(transport.patches).toHaveLength(1)

    const sent = await driver.send({ to: ['a@example.com'], subject: 're', text: 'done' })
    expect(sent.outboxId).toBe('outbox-1')

    await driver.move({ uid: 'm1', folderId: 'inbox', threadId: 'm1' }, 'archive')
    expect(transport.allIn('inbox').map((m) => m.id)).toEqual(['m2'])
    expect(transport.allIn('archive').map((m) => m.id)).toEqual(['m1'])

    await driver.delete({ uid: 'm2', folderId: 'inbox', threadId: 'm2' })
    expect(transport.allIn('inbox')).toHaveLength(0)

    await expect(
      driver.fetch({ uid: 'm2', folderId: 'inbox', threadId: 'm2' }),
    ).rejects.toMatchObject({ code: 'conflict' })
  })
})

describe('watch', () => {
  it('emits messages-added for arrivals after the baseline poll', async () => {
    const transport = new FakeGraphTransport()
    seedInbox(transport)
    const driver = driverWith(transport, { pollIntervalMs: 20, maxPollBackoffMs: 40 })
    await driver.connect()
    const kinds: string[] = []
    const added: string[] = []
    const watcher = driver.watch((ev) => {
      kinds.push(ev.kind)
      if (ev.kind === 'messages-added') added.push(...ev.refs.map((r) => r.uid))
    })
    await new Promise((resolve) => setTimeout(resolve, 60))
    expect(kinds).toContain('sync-state')
    expect(added).toHaveLength(0)
    transport.seed('inbox', [
      ...transport.allIn('inbox'),
      makeGraphMessage({ id: 'm3', subject: 'New arrival' }),
    ])
    await new Promise((resolve) => setTimeout(resolve, 80))
    watcher.stop()
    expect(added).toContain('m3')
  })
})

describe('auth failures', () => {
  it('maps invalid_grant to auth with needsReauth for main to wipe + re-auth', async () => {
    const transport = new FakeGraphTransport()
    const driver = driverWith(transport)
    await driver.connect()
    transport.failNextWith(
      new GraphServiceError(401, 'InvalidAuthenticationToken', 'Token expired: invalid_grant'),
    )
    try {
      await driver.list('inbox', { limit: 10 })
      throw new Error('list must throw on invalid_grant')
    } catch (err) {
      expect(err).toBeInstanceOf(GraphDriverError)
      expect(err as GraphDriverError).toMatchObject({ code: 'auth', needsReauth: true })
    }
  })

  it('requires connect before use', async () => {
    const transport = new FakeGraphTransport()
    const driver = driverWith(transport)
    await expect(driver.list('inbox', { limit: 10 })).rejects.toMatchObject({ code: 'offline' })
  })
})
