/**
 * BUG-1 regression (P4 sync): UIDVALIDITY full resync against the real
 * SQLite store must replace folder rows, not upsert over them.
 *
 * The engine calls SyncStorePort.purgeFolder on full resync; SqliteMailStore
 * used to expose only purgeFolderMessages, so the optional call was skipped
 * and ghost rows (still FTS-searchable) survived. Local-only, fakes only:
 * real SqliteMailStore on :memory: plus a scripted driver, no network.
 */
import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import type { MailEvent } from '@genoffice/mail-protocols'
import { SyncEngine } from '../src/main/sync/engine'
import { closeAllMailStores, SqliteMailStore } from '../src/main/store/index'
import { resetMailStoreConfig } from '../src/main/mail-store'
import { ScriptedDriver, syncState, testEnvelope } from './sync-fixtures'

describe('uidvalidity resync on the real SQLite store (BUG-1)', () => {
  let rootDir: string
  let store: SqliteMailStore

  beforeEach(() => {
    rootDir = mkdtempSync(join(tmpdir(), 'mail-resync-bug1-'))
    store = SqliteMailStore.open('a', { rootDir, dbPath: ':memory:' })
    store.upsertAccount({
      accountId: 'a',
      label: 'Test',
      kind: 'imap',
      sessionOnly: false,
      createdAt: '2026-10-14T09:00:00.000Z',
    })
  })

  afterEach(() => {
    store.close()
    closeAllMailStores()
    resetMailStoreConfig()
  })

  it('replaces folder rows, cleans FTS, and preserves queues', async () => {
    store.upsertEnvelopes([
      testEnvelope('old-1', { subject: 'Ghost old-1' }),
      testEnvelope('old-2', { subject: 'Ghost old-2' }),
    ])
    store.setSyncState(syncState('a', { uidValidity: 1, highestModseq: '10' }))

    const events: MailEvent[] = []
    const engine = new SyncEngine({
      store,
      noTimers: true,
      emit: {
        mailEvent: (event) => events.push(event),
        offlineStatus: () => undefined,
      },
    })
    const driver = new ScriptedDriver({ accountId: 'a', kind: 'imap' })
    engine.registerDriver(driver)

    // Queues that a resync must never drop.
    const outboxRecord = engine.enqueueSend('a', {
      to: ['reader@example.com'],
      subject: 'Queued while offline',
      text: 'hello',
    })
    store.saveDraft('a', 'thread-x', 'local draft text')
    store.queueIntent({
      accountId: 'a',
      kind: 'flag',
      targetUid: 'old-1',
      payloadJson: '{"seen":true,"folderId":"inbox"}',
    })

    driver.queueResult({
      decision: { kind: 'full-resync', reason: 'uidvalidity-changed' },
      envelopes: [
        testEnvelope('new-1', { subject: 'Fresh new-1' }),
        testEnvelope('new-2', { subject: 'Fresh new-2' }),
      ],
      state: syncState('a', { uidValidity: 2, highestModseq: '1' }),
    })

    const started = await engine.syncNow('a')
    expect(started).toEqual({ started: true })

    // Only fresh rows remain; ghost rows are gone from content and FTS.
    expect(store.getThread('thread-old-1')).toEqual([])
    expect(store.getThread('thread-old-2')).toEqual([])
    expect(store.getThread('thread-new-1').map((env) => env.subject)).toEqual(['Fresh new-1'])
    expect(store.fts('Ghost', ['a'])).toEqual([])
    expect(
      store
        .fts('Fresh', ['a'])
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
    expect(store.pendingIntents('a')).toHaveLength(1)
    expect(store.getSyncState('a', 'inbox')?.uidValidity).toBe(2)
  })
})
