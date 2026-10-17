/**
 * GmailDriver round-trip matrix row: the full MailProvider surface against
 * the in-memory fake transport. No network, no mailbox credentials.
 */

import { describe, expect, it } from 'vitest'
import { parseMimeMessage } from '../src/mime'
import {
  GmailDriver,
  buildGmailQuery,
  gmailLabelToFolder,
  gmailLabelsToFlags,
  parseSentRaw,
  type MailEvent,
} from '../src/gmail'
import { FakeGmailTransport, fakeAuth, memorySyncHooks, seedInbox } from './gmail-fake'

function driverWithInbox() {
  const transport = new FakeGmailTransport(seedInbox())
  const driver = new GmailDriver({
    auth: fakeAuth(),
    transport,
    makeOutboxId: () => 'outbox-1',
    pollMs: 60000,
    fromAddress: 'me@example.com',
  })
  return { driver, transport }
}

describe('gmail folders and envelopes', () => {
  it('maps system labels to roles and skips flag labels', async () => {
    const { driver } = driverWithInbox()
    await driver.connect()
    const folders = await driver.listFolders()
    const byId = new Map(folders.map((f) => [f.id, f]))
    expect(byId.get('INBOX')?.role).toBe('inbox')
    expect(byId.get('SENT')?.role).toBe('sent')
    expect(byId.get('DRAFT')?.role).toBe('drafts')
    expect(byId.get('TRASH')?.role).toBe('trash')
    expect(byId.get('SPAM')?.role).toBe('spam')
    expect(byId.get('Label_1')?.role).toBe('custom')
    expect(byId.has('UNREAD')).toBe(false)
    expect(byId.has('IMPORTANT')).toBe(false)
    expect(gmailLabelToFolder({ id: 'UNREAD' })).toBeUndefined()
    await driver.disconnect()
  })

  it('lists envelopes with flags, labels, and attachment marks', async () => {
    const { driver } = driverWithInbox()
    await driver.connect()
    const envelopes = await driver.list('INBOX', { limit: 10 })
    expect(envelopes).toHaveLength(2)
    const first = envelopes[0]
    expect(first.ref.uid).toBe('m1')
    expect(first.ref.threadId).toBe('t1')
    expect(first.subject).toBe('Invoice for October')
    expect(first.from).toEqual(['Alice Example <alice@example.com>'])
    expect(first.flags).toEqual({ seen: false, flagged: false, answered: false, draft: false })
    expect(first.hasAttach).toBe(true)
    expect(first.labels).toContain('INBOX')
    const second = envelopes[1]
    expect(second.flags).toEqual({ seen: true, flagged: true, answered: false, draft: false })
    expect(second.hasAttach).toBe(false)
    await driver.disconnect()
  })

  it('maps label sets to flags', () => {
    expect(gmailLabelsToFlags(['INBOX', 'UNREAD'])).toMatchObject({ seen: false, flagged: false })
    expect(gmailLabelsToFlags(['INBOX', 'STARRED'])).toMatchObject({ seen: true, flagged: true })
    expect(gmailLabelsToFlags(['INBOX', 'DRAFT'])).toMatchObject({ draft: true })
  })
})

describe('gmail search', () => {
  it('builds server queries from MailQuery', () => {
    expect(buildGmailQuery({ text: 'invoice' })).toBe('invoice')
    expect(buildGmailQuery({ from: 'alice@example.com', onlyUnread: true })).toBe(
      'from:alice@example.com is:unread',
    )
    expect(buildGmailQuery({ since: '2026-10-01T00:00:00Z' })).toBe('after:2026/10/01')
  })

  it('searches the fake server explicitly', async () => {
    const { driver } = driverWithInbox()
    await driver.connect()
    const bySender = await driver.search('INBOX', { from: 'alice@example.com' })
    expect(bySender.map((e) => e.ref.uid)).toEqual(['m1'])
    const unread = await driver.search('INBOX', { onlyUnread: true })
    expect(unread.map((e) => e.ref.uid)).toEqual(['m1'])
    const text = await driver.search('INBOX', { text: 'approved' })
    expect(text.map((e) => e.ref.uid)).toEqual(['m2'])
    await driver.disconnect()
  })
})

describe('gmail fetch and raw bodies', () => {
  it('fetches plain-text bodies', async () => {
    const { driver } = driverWithInbox()
    await driver.connect()
    const body = await driver.fetch({ uid: 'm1', folderId: 'INBOX', threadId: 't1' })
    expect(body.text).toContain('October invoice is attached')
    await driver.disconnect()
  })

  it('keeps HTML for the main layer through fetchRaw', async () => {
    const { driver } = driverWithInbox()
    await driver.connect()
    const raw = await driver.fetchRaw({ uid: 'm2', folderId: 'INBOX', threadId: 't1' })
    expect(raw.text).toContain('Looks good')
    expect(raw.html).toContain('<b>good</b>')
    const plain = await driver.fetchRaw({ uid: 'm1', folderId: 'INBOX', threadId: 't1' })
    expect(plain.html).toBeUndefined()
    await driver.disconnect()
  })
})

describe('gmail send round-trip', () => {
  it('sends MIME that parses back through the Lane A parser', async () => {
    const { driver, transport } = driverWithInbox()
    await driver.connect()
    const result = await driver.send({
      to: ['bob@example.com'],
      subject: 'Round trip',
      text: 'Hello Bob',
      html: '<p>Hello Bob</p>',
    })
    expect(result).toEqual({ outboxId: 'outbox-1' })
    expect(transport.sentRaw).toHaveLength(1)
    const raw = Buffer.from(
      transport.sentRaw[0].replace(/-/g, '+').replace(/_/g, '/'),
      'base64',
    ).toString('utf8')
    const back = await parseSentRaw(raw)
    expect(back.subject).toBe('Round trip')
    expect(back.text).toContain('Hello Bob')
    const parsed = await parseMimeMessage(raw)
    expect(parsed.to.map((a) => a.address)).toEqual(['bob@example.com'])
    expect(parsed.html).toContain('Hello Bob')
    await driver.disconnect()
  })
})

describe('gmail move, flags, delete', () => {
  it('moves by label add and remove', async () => {
    const { driver, transport } = driverWithInbox()
    await driver.connect()
    await driver.move({ uid: 'm1', folderId: 'INBOX', threadId: 't1' }, 'Label_1')
    expect(transport.modified).toEqual([{ id: 'm1', add: ['Label_1'], remove: ['INBOX'] }])
    await driver.disconnect()
  })

  it('maps seen and flagged to UNREAD and STARRED', async () => {
    const { driver, transport } = driverWithInbox()
    await driver.connect()
    await driver.setFlags({ uid: 'm1', folderId: 'INBOX', threadId: 't1' }, { seen: true })
    await driver.setFlags({ uid: 'm2', folderId: 'INBOX', threadId: 't1' }, { flagged: false })
    expect(transport.modified[0]).toEqual({ id: 'm1', add: [], remove: ['UNREAD'] })
    expect(transport.modified[1]).toEqual({ id: 'm2', add: [], remove: ['STARRED'] })
    await driver.disconnect()
  })

  it('deletes through trash', async () => {
    const { driver, transport } = driverWithInbox()
    await driver.connect()
    await driver.delete({ uid: 'm1', folderId: 'INBOX', threadId: 't1' })
    expect(transport.trashed).toEqual(['m1'])
    await driver.disconnect()
  })
})

describe('gmail watch', () => {
  it('emits history events and persists the new historyId', async () => {
    const { driver, transport } = driverWithInbox()
    const hooks = memorySyncHooks()
    hooks.save({
      accountId: 'gmail-test',
      folderId: 'INBOX',
      historyId: '9000',
      lastSyncAt: '2026-10-14T09:00:00Z',
    })
    transport.stagedHistory.added = [
      {
        id: 'm1',
        threadId: 't1',
        labelIds: ['INBOX', 'UNREAD'],
        snippet: 'new',
      },
    ]
    driver.setSyncHooks(hooks)
    await driver.connect()
    const events: MailEvent[] = []
    const sub = driver.watch((ev) => events.push(ev))
    await driver.pollOnce()
    sub.stop()
    expect(events.some((e) => 'kind' in e && e.kind === 'messages-added')).toBe(true)
    expect(events.some((e) => 'kind' in e && e.kind === 'sync-state' && e.state === 'idle')).toBe(
      true,
    )
    expect(hooks.load('gmail-test')?.historyId).toBe('9000')
    await driver.disconnect()
  })
})
