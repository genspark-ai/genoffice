/**
 * Single node:sqlite access point for the mail store (P4 lane T).
 *
 * This is the ONLY module in the repo allowed to import `node:sqlite`.
 * Every other store file receives an opened DatabaseSync handle and never
 * touches the driver directly. (A repo-wide no-restricted-imports lint rule
 * should pin that; adding it touches shared config outside this lane, so it
 * is left for the Tech Lead — see the lane handoff.)
 *
 * Write discipline per the sync spec: WAL mode, a 5 s busy timeout, foreign
 * keys on, and one writer via the main-process handler queue (no worker
 * threads in P4). Migrations run through runMigrations() only.
 */

import { DatabaseSync } from 'node:sqlite'
import { MIGRATIONS, LATEST_STORE_VERSION } from './migrations'

export type { DatabaseSync }

export function openDatabase(path: string): DatabaseSync {
  const db = new DatabaseSync(path)
  db.exec('PRAGMA journal_mode = WAL')
  db.exec('PRAGMA busy_timeout = 5000')
  db.exec('PRAGMA foreign_keys = ON')
  return db
}

export function closeDatabase(db: DatabaseSync): void {
  db.close()
}

export function getUserVersion(db: DatabaseSync): number {
  const row = db.prepare('PRAGMA user_version').get() as unknown as Record<string, unknown>
  return typeof row?.['user_version'] === 'number' ? (row['user_version'] as number) : 0
}

function setUserVersion(db: DatabaseSync, version: number): void {
  db.exec(`PRAGMA user_version = ${version}`)
}

/** Run a callback inside one transaction; rolls back and rethrows on error. */
export function withTransaction<T>(db: DatabaseSync, fn: () => T): T {
  db.exec('BEGIN IMMEDIATE')
  try {
    const result = fn()
    db.exec('COMMIT')
    return result
  } catch (err) {
    try {
      db.exec('ROLLBACK')
    } catch {
      // Rollback of a failed transaction is best-effort; the original error wins.
    }
    throw err
  }
}

/**
 * Apply every pending migration forward-only, each in its own transaction.
 * Safe to call on every open; a current database is a no-op.
 */
export function runMigrations(db: DatabaseSync): void {
  const current = getUserVersion(db)
  for (const migration of MIGRATIONS) {
    if (migration.version <= current) continue
    withTransaction(db, () => {
      db.exec(migration.sql)
      setUserVersion(db, migration.version)
    })
  }
}

export function assertCurrentVersion(db: DatabaseSync): void {
  const current = getUserVersion(db)
  if (current !== LATEST_STORE_VERSION) {
    throw new Error(`mail store schema version ${current}, expected ${LATEST_STORE_VERSION}`)
  }
}
