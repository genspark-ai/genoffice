import { describe, expect, it } from 'vitest'
import { applyLocalSearch, FakeMailApi } from '../src/renderer/devFakes'
import {
  allowSender,
  isSenderAllowed,
  loadImageAllowlist,
  normalizeSender,
} from '../src/renderer/reader/allowlist'

/**
 * Lane R IPC-fake tests: paging, matchesQuery post-filter agreement,
 * mutations, outbox, server expansion, and the sender allowlist.
 */

describe('FakeMailApi reads', () => {
  it('lists accounts, folders, and pages messages with a cursor', async () => {
    const api = new FakeMailApi()
    const accounts = await api.listAccounts()
    expect(accounts.length).toBe(1)
    const folders = await api.listFolders(accounts[0].accountId)
    expect(folders.map((folder) => folder.id)).toContain('inbox')
    const first = await api.listMessages(accounts[0].accountId, 'inbox', 2)
    expect(first.envelopes.length).toBe(2)
    expect(first.cursor).toBeDefined()
    const second = await api.listMessages(accounts[0].accountId, 'inbox', 10, first.cursor)
    expect(second.envelopes.length).toBeGreaterThan(0)
    expect(second.cursor).toBeUndefined()
    const uids = new Set([...first.envelopes, ...second.envelopes].map((env) => env.ref.uid))
    expect(uids.size).toBe(first.envelopes.length + second.envelopes.length)
  })

  it('searchLocal agrees with the matchesQuery post-filter', async () => {
    const api = new FakeMailApi()
    const hits = await api.searchLocal(['demo'], { text: 'invoice' })
    expect(hits.length).toBeGreaterThan(0)
    const page = await api.listMessages('demo', 'inbox', 50)
    expect(applyLocalSearch(page.envelopes, 'invoice')).toEqual(
      hits.filter((env) => env.ref.folderId === 'inbox'),
    )
    expect(applyLocalSearch(page.envelopes, '   ')).toEqual(page.envelopes)
  })

  it('searchServer surfaces server-only extras', async () => {
    const api = new FakeMailApi()
    const remote = await api.searchServer('demo', 'inbox', { text: 'invoice' })
    expect(remote.some((env) => env.ref.uid === 'srv1')).toBe(true)
  })

  it('fetchBody returns seeded text and fake-inline HTML', async () => {
    const api = new FakeMailApi()
    const body = await api.fetchBody({ uid: 'm1', folderId: 'inbox', threadId: 'thread-m1' })
    expect(body.text).toContain('INV-1042')
    expect((body as { html?: string }).html).toContain('INV-1042')
  })
})

describe('FakeMailApi writes', () => {
  it('moves, flags, and deletes envelopes', async () => {
    const api = new FakeMailApi()
    await api.setFlags({ uid: 'm1', folderId: 'inbox', threadId: 'thread-m1' }, { seen: true })
    const afterFlags = await api.listMessages('demo', 'inbox', 50)
    expect(afterFlags.envelopes.find((env) => env.ref.uid === 'm1')?.flags.seen).toBe(true)
    await api.move({ uid: 'm1', folderId: 'inbox', threadId: 'thread-m1' }, 'archive')
    const afterMove = await api.listMessages('demo', 'inbox', 50)
    expect(afterMove.envelopes.some((env) => env.ref.uid === 'm1')).toBe(false)
    await api.deleteMessage({ uid: 'm1', folderId: 'archive', threadId: 'thread-m1' })
    const archived = await api.listMessages('demo', 'archive', 50)
    expect(archived.envelopes.some((env) => env.ref.uid === 'm1')).toBe(false)
  })

  it('queues sends and reports outbox status', async () => {
    const api = new FakeMailApi()
    const queued = await api.outboxStatus('demo')
    expect(queued).toEqual({ queued: 0, failed: 0 })
    const sent = await api.composeSend({ to: ['a@example.com'], subject: 'hi', text: 'hello' })
    expect(sent.outboxId).toContain('fake-outbox-')
    expect((await api.outboxStatus('demo')).queued).toBe(1)
    await api.saveDraft('demo', 'thread-x', 'draft text')
  })

  it('exposes fake phishing sidecars for seeded messages', async () => {
    const api = new FakeMailApi()
    expect(api.fakeSecurity({ uid: 'm3', folderId: 'inbox', threadId: 'thread-m3' }).replyTo).toBe(
      'outside@example.org',
    )
    expect(api.fakeSecurity({ uid: 'm1', folderId: 'inbox', threadId: 'thread-m1' })).toEqual({})
  })
})

describe('sender allowlist', () => {
  it('normalizes display-name addresses and round-trips entries', () => {
    expect(normalizeSender('Acme <Billing@Example.COM>')).toBe('billing@example.com')
    const before = new Set(loadImageAllowlist())
    allowSender('Test Sender <allow-me@example.com>')
    expect(isSenderAllowed('allow-me@example.com')).toBe(true)
    expect(isSenderAllowed('other@example.com')).toBe(false)
    expect(new Set(loadImageAllowlist()).has('allow-me@example.com')).toBe(true)
    expect(before.has('allow-me@example.com')).toBe(false)
  })
})
