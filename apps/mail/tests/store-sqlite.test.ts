import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import type { MailEnvelope } from '@genoffice/mail-protocols'
import {
  INLINE_BODY_CAP_BYTES,
  LATEST_STORE_VERSION,
  OUTBOX_MAX_ATTEMPTS,
  SqliteMailStore,
  closeAllMailStores,
  nextOutboxRetryDelay,
  sanitizeErrorText,
} from '../src/main/store/index'
import { configureMailStoreDir, getMailStore, resetMailStoreConfig } from '../src/main/mail-store'

function envelope(uid: string, subject: string, folderId = 'inbox'): MailEnvelope {
  return {
    ref: { uid, folderId, threadId: 'thread-1' },
    from: ['sender@example.com'],
    to: ['reader@example.com'],
    cc: [],
    subject,
    date: '2026-10-14T10:00:00.000Z',
    flags: { seen: false, flagged: false, answered: false, draft: false },
    labels: [],
    snippet: `snippet for ${subject}`,
    hasAttach: false,
  }
}

describe('SqliteMailStore schema', () => {
  let rootDir: string
  let store: SqliteMailStore

  beforeEach(() => {
    rootDir = mkdtempSync(join(tmpdir(), 'mail-store-'))
    store = SqliteMailStore.open('acct-1', { rootDir, dbPath: ':memory:' })
  })

  afterEach(() => {
    store.close()
    closeAllMailStores()
    resetMailStoreConfig()
  })

  it('migrates a fresh database to the latest version', () => {
    expect(store.schemaVersion).toBe(LATEST_STORE_VERSION)
  })

  it('round-trips accounts with an opaque credential reference', () => {
    store.upsertAccount({
      accountId: 'acct-1',
      label: 'Work',
      kind: 'imap',
      sessionOnly: false,
      credRef: 'vault-handle-123',
      createdAt: '2026-10-14T09:00:00.000Z',
    })
    const account = store.getAccount('acct-1')
    expect(account?.label).toBe('Work')
    expect(account?.kind).toBe('imap')
    expect(account?.credRef).toBe('vault-handle-123')
    expect(store.getAccount('other-acct')).toBeUndefined()
  })

  it('upserts and lists folders, and cascades messages on folder delete', () => {
    store.upsertFolder('acct-1', { id: 'inbox', name: 'Inbox', role: 'inbox', unread: 2 })
    store.upsertFolder('acct-1', { id: 'inbox', name: 'Inbox', role: 'inbox', unread: 5 })
    expect(store.listFolders('acct-1')).toHaveLength(1)
    expect(store.listFolders('acct-1')[0].unread).toBe(5)
    store.setFolderUnread('acct-1', 'inbox', 1)
    expect(store.listFolders('acct-1')[0].unread).toBe(1)
    store.upsertEnvelopes([envelope('u1', 'Hello')])
    store.deleteFolder('acct-1', 'inbox')
    expect(store.listFolders('acct-1')).toHaveLength(0)
    expect(store.getThread('thread-1')).toHaveLength(0)
  })
})

describe('SqliteMailStore MailStore compatibility', () => {
  let rootDir: string
  let store: SqliteMailStore

  beforeEach(() => {
    rootDir = mkdtempSync(join(tmpdir(), 'mail-store-'))
    store = SqliteMailStore.open('acct-1', { rootDir, dbPath: ':memory:' })
  })

  afterEach(() => {
    store.close()
    closeAllMailStores()
    resetMailStoreConfig()
  })

  it('stores envelopes and serves threads newest-agnostic', () => {
    store.upsertEnvelopes([envelope('u1', 'Invoice'), envelope('u2', 'Lunch')])
    expect(store.getThread('thread-1')).toHaveLength(2)
    const page = store.listEnvelopes('acct-1', 'inbox', { limit: 1 })
    expect(page).toHaveLength(1)
  })

  it('round-trips drafts', () => {
    expect(store.draftFor('acct-1', 't1')).toBeUndefined()
    store.saveDraft('acct-1', 't1', 'draft text')
    expect(store.draftFor('acct-1', 't1')).toBe('draft text')
    store.deleteDraft('acct-1', 't1')
    expect(store.draftFor('acct-1', 't1')).toBeUndefined()
  })

  it('queues and resolves offline intents', () => {
    const id = store.queueIntent({
      accountId: 'acct-1',
      kind: 'move',
      targetUid: 'u1',
      payloadJson: '{"toFolderId":"archive"}',
    })
    expect(store.pendingIntents('acct-1')).toHaveLength(1)
    store.bumpIntentAttempt(id)
    expect(store.listIntents('acct-1')[0].attempts).toBe(1)
    store.markIntent(id, 'acked')
    expect(store.pendingIntents('acct-1')).toHaveLength(0)
  })

  it('round-trips sync state for every driver shape', () => {
    expect(store.getSyncState('acct-1', 'inbox')).toBeUndefined()
    store.setSyncState({
      accountId: 'acct-1',
      folderId: 'inbox',
      uidValidity: 42,
      highestModseq: '900',
      lastSyncAt: '2026-10-14T10:00:00.000Z',
    })
    expect(store.getSyncState('acct-1', 'inbox')?.uidValidity).toBe(42)
    store.setSyncState({
      accountId: 'acct-1',
      folderId: 'inbox',
      historyId: 'h99',
      deltaToken: 'd99',
      lastSyncAt: '2026-10-14T11:00:00.000Z',
    })
    const state = store.getSyncState('acct-1', 'inbox')
    expect(state?.historyId).toBe('h99')
    expect(state?.deltaToken).toBe('d99')
    expect(state?.uidValidity).toBeUndefined()
  })

  it('never leaks rows across accounts', () => {
    store.upsertEnvelopes([envelope('u1', 'Private')])
    expect(store.getEnvelope('other', 'inbox', 'u1')).toBeUndefined()
    expect(store.listEnvelopes('other', 'inbox')).toEqual([])
    expect(store.fts('Private', ['other'])).toEqual([])
    expect(store.pendingIntents('other')).toEqual([])
    expect(store.draftFor('other', 't1')).toBeUndefined()
  })
})

describe('SqliteMailStore full message writes', () => {
  let rootDir: string
  let store: SqliteMailStore

  beforeEach(() => {
    rootDir = mkdtempSync(join(tmpdir(), 'mail-store-'))
    store = SqliteMailStore.open('acct-1', { rootDir, dbPath: ':memory:' })
  })

  afterEach(() => {
    store.close()
    closeAllMailStores()
    resetMailStoreConfig()
  })

  it('stores plaintext inline and HTML content-addressed on disk', () => {
    store.upsertMessageFull({
      folderId: 'inbox',
      uid: 'u1',
      threadId: 'thread-1',
      from: ['sender@example.com'],
      to: ['reader@example.com'],
      subject: 'Full body',
      date: '2026-10-14T10:00:00.000Z',
      bodyText: 'Hello reader, see the invoice attached.',
      bodyHtml: '<p>Hello <b>reader</b>, see the invoice attached.</p>',
    })
    const body = store.getMessageBody('acct-1', 'inbox', 'u1')
    expect(body?.text).toBe('Hello reader, see the invoice attached.')
    expect(body?.truncated).toBe(false)
    expect(body?.htmlRef).toMatch(/^blobs\/[0-9a-f]{2}\/[0-9a-f]{64}\.html$/)
    expect(store.readBodyHtml(body?.htmlRef ?? '')).toContain('<b>reader</b>')
    expect(store.getEnvelope('acct-1', 'inbox', 'u1')?.snippet).toContain('Hello reader')
  })

  it('derives the snippet from HTML when no plaintext exists', () => {
    store.upsertMessageFull({
      folderId: 'inbox',
      uid: 'u2',
      threadId: 'thread-1',
      from: ['sender@example.com'],
      to: ['reader@example.com'],
      subject: 'HTML only',
      date: '2026-10-14T10:00:00.000Z',
      bodyHtml: '<p>Quarterly <i>results</i> are in.</p>',
    })
    expect(store.getEnvelope('acct-1', 'inbox', 'u2')?.snippet).toContain('Quarterly')
  })

  it('flags over-cap plaintext instead of silently dropping it', () => {
    const big = `x`.repeat(INLINE_BODY_CAP_BYTES + 10)
    store.upsertMessageFull({
      folderId: 'inbox',
      uid: 'u3',
      threadId: 'thread-1',
      from: ['sender@example.com'],
      to: ['reader@example.com'],
      subject: 'Big',
      date: '2026-10-14T10:00:00.000Z',
      bodyText: big,
    })
    const body = store.getMessageBody('acct-1', 'inbox', 'u3')
    expect(body?.text.length).toBe(INLINE_BODY_CAP_BYTES)
    expect(body?.truncated).toBe(true)
  })

  it('keeps stored bodies when a later envelope-only upsert arrives', () => {
    store.upsertMessageFull({
      folderId: 'inbox',
      uid: 'u4',
      threadId: 'thread-1',
      from: ['sender@example.com'],
      to: ['reader@example.com'],
      subject: 'Bodied',
      date: '2026-10-14T10:00:00.000Z',
      bodyText: 'kept body',
    })
    store.upsertEnvelopes([envelope('u4', 'Bodied')])
    expect(store.getMessageBody('acct-1', 'inbox', 'u4')?.text).toBe('kept body')
  })

  it('moves, flags, labels, and deletes messages', () => {
    store.upsertEnvelopes([envelope('u5', 'Todo')])
    store.setMessageFlags('acct-1', 'inbox', 'u5', { seen: true, flagged: true })
    expect(store.getEnvelope('acct-1', 'inbox', 'u5')?.flags.seen).toBe(true)
    store.setMessageLabels('acct-1', 'inbox', 'u5', ['Starred'])
    expect(store.getEnvelope('acct-1', 'inbox', 'u5')?.labels).toEqual(['Starred'])
    store.setMessageModseq('acct-1', 'inbox', 'u5', '901')
    store.moveMessage('acct-1', { uid: 'u5', folderId: 'inbox', threadId: 'thread-1' }, 'archive')
    expect(store.getEnvelope('acct-1', 'archive', 'u5')?.subject).toBe('Todo')
    expect(store.getEnvelope('acct-1', 'inbox', 'u5')).toBeUndefined()
    store.deleteMessage('acct-1', 'archive', 'u5')
    expect(store.getEnvelope('acct-1', 'archive', 'u5')).toBeUndefined()
  })
})

describe('SqliteMailStore FTS', () => {
  let rootDir: string
  let store: SqliteMailStore

  beforeEach(() => {
    rootDir = mkdtempSync(join(tmpdir(), 'mail-store-'))
    store = SqliteMailStore.open('acct-1', { rootDir, dbPath: ':memory:' })
    store.upsertMessageFull({
      folderId: 'inbox',
      uid: 'u1',
      threadId: 'thread-1',
      from: ['accounts@example.com'],
      to: ['reader@example.com'],
      subject: 'September invoice total',
      date: '2026-10-14T10:00:00.000Z',
      bodyText: 'The quarterly total is due next Friday.',
      bodyHtml: '<p>Pay the <b>secret-html-token</b> now.</p>',
    })
  })

  afterEach(() => {
    store.close()
    closeAllMailStores()
    resetMailStoreConfig()
  })

  it('searches subject, sender, snippet, and plaintext', () => {
    expect(store.searchLocal('invoice').map((env) => env.ref.uid)).toEqual(['u1'])
    expect(store.searchLocal('accounts@example').map((env) => env.ref.uid)).toEqual(['u1'])
    expect(store.searchLocal('quarterly').map((env) => env.ref.uid)).toEqual(['u1'])
    expect(store.fts('invoice', ['acct-1'])).toHaveLength(1)
  })

  it('never indexes raw HTML blobs', () => {
    expect(store.searchLocal('secret-html-token')).toEqual([])
  })

  it('retrieves a CJK subject from a short query', () => {
    store.upsertMessageFull({
      folderId: 'inbox',
      uid: 'cjk-1',
      threadId: 'thread-2',
      from: ['sender@example.com'],
      to: ['reader@example.com'],
      subject: '邮件测试报告',
      date: '2026-10-14T10:00:00.000Z',
      bodyText: 'plain body',
    })
    expect(store.searchLocal('邮件').map((env) => env.ref.uid)).toEqual(['cjk-1'])
    expect(store.searchLocal('测试报告').map((env) => env.ref.uid)).toEqual(['cjk-1'])
  })

  it('treats user syntax as literal text, never MATCH operators', () => {
    expect(store.searchLocal('"invoice" OR 1=1')).toEqual([])
    expect(store.searchLocal('   ')).toEqual([])
  })

  it('mirrors updates and deletes into the FTS index', () => {
    store.upsertMessageFull({
      folderId: 'inbox',
      uid: 'u1',
      threadId: 'thread-1',
      from: ['accounts@example.com'],
      to: ['reader@example.com'],
      subject: 'October receipt total',
      date: '2026-10-14T10:00:00.000Z',
      bodyText: 'The quarterly total is due next Friday.',
    })
    expect(store.searchLocal('invoice')).toEqual([])
    expect(store.searchLocal('receipt').map((env) => env.ref.uid)).toEqual(['u1'])
    store.deleteMessage('acct-1', 'inbox', 'u1')
    expect(store.searchLocal('receipt')).toEqual([])
  })

  it('rebuilds the FTS mirror on demand', () => {
    store.rebuildFts()
    expect(store.searchLocal('invoice').map((env) => env.ref.uid)).toEqual(['u1'])
  })
})

describe('SqliteMailStore outbox', () => {
  let rootDir: string
  let store: SqliteMailStore
  let tick = 0

  beforeEach(() => {
    rootDir = mkdtempSync(join(tmpdir(), 'mail-store-'))
    tick = 0
    store = SqliteMailStore.open('acct-1', {
      rootDir,
      dbPath: ':memory:',
      clock: () => `2026-10-14T10:00:${String(tick++).padStart(2, '0')}.000Z`,
    })
  })

  afterEach(() => {
    store.close()
    closeAllMailStores()
    resetMailStoreConfig()
  })

  function req(subject: string) {
    return { to: ['reader@example.com'], subject, text: 'body' }
  }

  it('sends FIFO with one in-flight claim at a time', () => {
    const first = store.enqueueOutbox('acct-1', req('first'))
    const second = store.enqueueOutbox('acct-1', req('second'))
    expect(first).not.toBe(second)
    const claimed = store.claimNextOutbox('acct-1')
    expect(claimed?.id).toBe(first)
    expect(claimed?.state).toBe('sending')
    expect(claimed?.attempts).toBe(1)
    expect(store.claimNextOutbox('acct-1')?.id).toBe(second)
    expect(store.claimNextOutbox('acct-1')).toBeUndefined()
    expect(store.outboxCounts('acct-1')).toEqual({ queued: 0, failed: 0 })
  })

  it('marks sends and prunes them after the retention window', () => {
    const id = store.enqueueOutbox('acct-1', req('hello'))
    store.claimNextOutbox('acct-1')
    store.markOutboxSent(id, '2026-10-14T10:05:00.000Z')
    expect(store.getOutbox(id)?.state).toBe('sent')
    expect(store.pruneSentOutbox('acct-1', '2026-10-21T10:05:01.000Z')).toBe(1)
    expect(store.getOutbox(id)).toBeUndefined()
  })

  it('fails with sanitized errors and supports manual retry', () => {
    const id = store.enqueueOutbox('acct-1', req('hello'))
    store.claimNextOutbox('acct-1')
    store.failOutbox(
      id,
      'SMTP 550\nmailbox unavailable\nto: reader@example.com',
      '2026-10-14T10:00:05.000Z',
    )
    const failed = store.getOutbox(id)
    expect(failed?.state).toBe('failed')
    expect(failed?.lastError).not.toContain('\n')
    expect(store.outboxCounts('acct-1')).toEqual({ queued: 0, failed: 1 })
    store.retryOutbox(id)
    const retried = store.getOutbox(id)
    expect(retried?.state).toBe('queued')
    expect(retried?.attempts).toBe(0)
  })

  it('cancels queued entries but never sent or sending ones', () => {
    const queued = store.enqueueOutbox('acct-1', req('queued'))
    expect(store.cancelOutbox(queued)).toBe(true)
    expect(store.getOutbox(queued)?.state).toBe('cancelled')
    const sending = store.enqueueOutbox('acct-1', req('sending'))
    store.claimNextOutbox('acct-1')
    expect(store.cancelOutbox(sending)).toBe(false)
    expect(store.getOutbox(sending)?.state).toBe('sending')
  })

  it('lists newest-first within the spec cap and deletes entries', () => {
    for (let index = 0; index < 5; index += 1) store.enqueueOutbox('acct-1', req(`m${index}`))
    const listed = store.listOutbox('acct-1', { limit: 3 })
    expect(listed).toHaveLength(3)
    expect(listed[0].req.subject).toBe('m4')
    store.deleteOutbox(listed[0].id)
    expect(store.listOutbox('acct-1')).toHaveLength(4)
  })

  it('keeps attachment memory payloads out of the persisted request', () => {
    const id = store.enqueueOutbox('acct-1', {
      to: ['reader@example.com'],
      subject: 'with file',
      text: 'see attached',
      attachments: [
        {
          filename: 'a.bin',
          mime: 'application/octet-stream',
          size: 3,
          path: '/tmp/a.bin',
          data: new Uint8Array([1, 2, 3]),
        },
      ],
    })
    expect(store.getOutbox(id)?.req.attachments?.[0]).toEqual({
      filename: 'a.bin',
      mime: 'application/octet-stream',
      size: 3,
      path: '/tmp/a.bin',
    })
  })

  it('exposes the locked retry schedule', () => {
    expect([
      nextOutboxRetryDelay(1),
      nextOutboxRetryDelay(2),
      nextOutboxRetryDelay(3),
      nextOutboxRetryDelay(4),
    ]).toEqual([5_000, 30_000, 120_000, 600_000])
    expect(nextOutboxRetryDelay(0)).toBeNull()
    expect(nextOutboxRetryDelay(OUTBOX_MAX_ATTEMPTS)).toBeNull()
  })

  it('sanitizes failure text to one capped line', () => {
    expect(sanitizeErrorText('a\nb\r\nc')).toBe('a b c')
    expect(sanitizeErrorText('x'.repeat(10_000)).length).toBeLessThanOrEqual(500)
  })
})

describe('SqliteMailStore attachments', () => {
  let rootDir: string
  let store: SqliteMailStore

  beforeEach(() => {
    rootDir = mkdtempSync(join(tmpdir(), 'mail-store-'))
    store = SqliteMailStore.open('acct-1', { rootDir, dbPath: ':memory:' })
  })

  afterEach(() => {
    store.close()
    closeAllMailStores()
    resetMailStoreConfig()
  })

  it('walks staged to quarantined to stored', () => {
    store.addAttachment({
      accountId: 'acct-1',
      msgUid: 'u1',
      filename: 'report.pdf',
      mime: 'application/pdf',
      size: 12,
      blobPath: join(rootDir, 'staging', 'tmp-report.pdf'),
      state: 'staged',
    })
    store.setAttachmentState('acct-1', 'u1', 'report.pdf', 'quarantined')
    store.setAttachmentState('acct-1', 'u1', 'report.pdf', 'stored', 'attachments/ab/report.pdf')
    expect(store.listAttachments('acct-1', 'u1')[0].state).toBe('stored')
    store.deleteAttachment('acct-1', 'u1', 'report.pdf')
    expect(store.listAttachments('acct-1', 'u1')).toEqual([])
  })

  it('rejects stored blobs outside the store tree', () => {
    store.addAttachment({
      accountId: 'acct-1',
      msgUid: 'u1',
      filename: 'evil.pdf',
      mime: 'application/pdf',
      size: 1,
      blobPath: join(rootDir, 'staging', 'evil.pdf'),
      state: 'staged',
    })
    expect(() =>
      store.setAttachmentState('acct-1', 'u1', 'evil.pdf', 'stored', '/tmp/evil.pdf'),
    ).toThrow()
  })
})

describe('SqliteMailStore resync purge', () => {
  let rootDir: string
  let store: SqliteMailStore

  beforeEach(() => {
    rootDir = mkdtempSync(join(tmpdir(), 'mail-store-'))
    store = SqliteMailStore.open('acct-1', { rootDir, dbPath: ':memory:' })
  })

  afterEach(() => {
    store.close()
    closeAllMailStores()
    resetMailStoreConfig()
  })

  it('drops folder rows but keeps outbox, drafts, and intents', () => {
    store.upsertEnvelopes([envelope('u1', 'Old')])
    const outboxId = store.enqueueSend({ to: ['b@example.com'], subject: 's', text: 'b' })
    store.saveDraft('acct-1', 't1', 'local draft wins')
    const intentId = store.queueIntent({
      accountId: 'acct-1',
      kind: 'flag',
      targetUid: 'u1',
      payloadJson: '{"seen":true}',
    })
    expect(store.purgeFolderMessages('acct-1', 'inbox')).toBe(1)
    expect(store.getThread('thread-1')).toEqual([])
    expect(store.searchLocal('Old')).toEqual([])
    expect(store.getOutbox(outboxId)?.state).toBe('queued')
    expect(store.draftFor('acct-1', 't1')).toBe('local draft wins')
    expect(store.pendingIntents('acct-1').map((intent) => intent.id)).toEqual([intentId])
  })
})

describe('getMailStore cutover', () => {
  let rootDir: string

  beforeEach(() => {
    rootDir = mkdtempSync(join(tmpdir(), 'mail-store-'))
  })

  afterEach(() => {
    resetMailStoreConfig()
  })

  it('serves the in-memory stub until a store dir is configured', () => {
    getMailStore().upsertEnvelopes([
      { ...envelope('u1', 'Stub'), ref: { uid: 'u1', folderId: 'inbox', threadId: 'thread-stub' } },
    ])
    expect(getMailStore().getThread('thread-stub')).toHaveLength(1)
  })

  it('resolves per-account SQLite stores behind the same accessor', () => {
    configureMailStoreDir(rootDir)
    const sqlite = getMailStore('acct-9')
    sqlite.upsertEnvelopes([
      {
        ...envelope('u1', 'Persisted'),
        ref: { uid: 'u1', folderId: 'inbox', threadId: 'thread-cutover' },
      },
    ])
    expect(sqlite.getThread('thread-cutover')).toHaveLength(1)
    expect(getMailStore().getThread('thread-cutover')).toHaveLength(0)
    closeAllMailStores()
    expect(getMailStore('acct-9').getThread('thread-cutover')).toHaveLength(1)
  })
})
