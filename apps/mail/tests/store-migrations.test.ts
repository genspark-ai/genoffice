import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import {
  LATEST_STORE_VERSION,
  MIGRATIONS,
  SqliteMailStore,
  closeAllMailStores,
  closeDatabase,
  getUserVersion,
  openDatabase,
} from '../src/main/store/index'
import type { DatabaseSync } from '../src/main/store/sqlite'

afterEach(() => {
  closeAllMailStores()
})

/** Build a version-1 database the way an older client would have left it. */
function buildV1Database(file: string): void {
  const db: DatabaseSync = openDatabase(file)
  try {
    const first = MIGRATIONS[0]
    db.exec(first.sql)
    db.exec(`INSERT INTO accounts (accountId, label, kind, sessionOnly, createdAt)
             VALUES ('acct-1', 'Work', 'imap', 0, '2026-10-01T00:00:00.000Z')`)
    db.exec(`INSERT INTO folders (accountId, folderId, name, role, unread)
             VALUES ('acct-1', 'inbox', 'Inbox', 'inbox', 1)`)
    db.exec(`INSERT INTO messages
               (accountId, folderId, uid, threadId, fromJson, toJson, subject,
                dateIso, snippet, bodyText, updatedAt)
             VALUES ('acct-1', 'inbox', 'legacy-1', 'thread-legacy',
                     '["legacy@example.com"]', '["reader@example.com"]',
                     'Legacy migration row', '2026-10-01T00:00:00.000Z',
                     'legacy snippet', 'legacy body', '2026-10-01T00:00:00.000Z')`)
    db.exec('PRAGMA user_version = 1')
    expect(getUserVersion(db)).toBe(1)
  } finally {
    closeDatabase(db)
  }
}

describe('store migrations', () => {
  it('opens fresh databases at the latest version', () => {
    const rootDir = mkdtempSync(join(tmpdir(), 'mail-migrate-'))
    const store = SqliteMailStore.open('acct-1', { rootDir, dbPath: ':memory:' })
    try {
      expect(store.schemaVersion).toBe(LATEST_STORE_VERSION)
      expect(LATEST_STORE_VERSION).toBeGreaterThanOrEqual(2)
    } finally {
      store.close()
    }
  })

  it('upgrades a version-1 fixture to the latest schema without data loss', () => {
    const rootDir = mkdtempSync(join(tmpdir(), 'mail-migrate-'))
    const file = join(rootDir, 'legacy.sqlite')
    buildV1Database(file)

    const store = SqliteMailStore.open('acct-1', { rootDir, dbPath: file })
    try {
      expect(store.schemaVersion).toBe(LATEST_STORE_VERSION)
      const found = store.getEnvelope('acct-1', 'inbox', 'legacy-1')
      expect(found?.subject).toBe('Legacy migration row')
      expect(found?.snippet).toBe('legacy snippet')
      // The T-local bodyTruncated column defaults old rows to intact.
      expect(store.getMessageBody('acct-1', 'inbox', 'legacy-1')?.truncated).toBe(false)
      // The FTS mirror survives the upgrade and stays usable for new writes.
      expect(store.searchLocal('Legacy').map((env) => env.ref.uid)).toEqual(['legacy-1'])
      store.upsertEnvelopes([
        {
          ref: { uid: 'post-upgrade', folderId: 'inbox', threadId: 'thread-legacy' },
          from: ['new@example.com'],
          to: ['reader@example.com'],
          cc: [],
          subject: 'After upgrade',
          date: '2026-10-14T10:00:00.000Z',
          flags: { seen: false, flagged: false, answered: false, draft: false },
          labels: [],
          snippet: 'new row',
          hasAttach: false,
        },
      ])
      expect(store.searchLocal('After').map((env) => env.ref.uid)).toEqual(['post-upgrade'])
    } finally {
      store.close()
    }
  })

  it('re-running migrations on a current database is a no-op', () => {
    const rootDir = mkdtempSync(join(tmpdir(), 'mail-migrate-'))
    const file = join(rootDir, 'current.sqlite')
    buildV1Database(file)
    const first = SqliteMailStore.open('acct-1', { rootDir, dbPath: file })
    first.close()
    const second = SqliteMailStore.open('acct-1', { rootDir, dbPath: file })
    try {
      expect(second.schemaVersion).toBe(LATEST_STORE_VERSION)
      expect(second.getEnvelope('acct-1', 'inbox', 'legacy-1')?.subject).toBe(
        'Legacy migration row',
      )
    } finally {
      second.close()
    }
  })
})
