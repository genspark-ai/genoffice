/**
 * SQLite-backed MailStore for GenOffice Mail (P4 lane T).
 *
 * Per-account databases: one `<accountId>.sqlite` file per account under a
 * profile root (default `~/.genoffice/mail`), so removing an account is one
 * file delete and no query can leak rows across accounts. Full HTML bodies
 * live content-addressed on disk (`blobs/xx/<sha256>.html`); attachments
 * move staged-temp to quarantine to store, never as SQLite BLOBs. Plaintext
 * plus snippet stay inline in `messages`. Secrets never enter this store:
 * accounts carry only an opaque safeStorage `credRef`, and no log or error
 * path here emits bodies, addresses, or secrets (PII stays out of logs).
 *
 * Y lane entry points: openMailStore / closeMailStore, SqliteMailStore
 * (implements the MailStore interface, so getMailStore() can hand it out),
 * nextOutboxRetryDelay, sanitizeErrorText, OUTBOX_* constants.
 */

import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { randomUUID } from 'node:crypto'
import { join, resolve, sep } from 'node:path'
import { contentAddress, makeSnippet, stripTags } from '@genoffice/mail-protocols'
import type {
  MailEnvelope,
  MailFlags,
  MailFolder,
  MailIntent,
  MailMessageRef,
  MailStore,
  SendRequest,
  SyncState,
} from '@genoffice/mail-protocols'
import {
  closeDatabase,
  getUserVersion,
  openDatabase,
  runMigrations,
  withTransaction,
} from './sqlite'
import type { DatabaseSync } from './sqlite'

/** Inline plaintext cap: bodies over this keep the head plus a flag. */
export const INLINE_BODY_CAP_BYTES = 1024 * 1024
/** Snippet length: first characters of de-HTMLed text, per the sync spec. */
export const SNIPPET_CHARS = 280
/** Stored lastError cap: sanitized provider text, never secrets or PII. */
export const LAST_ERROR_CHARS = 500
/** Sent outbox rows are kept this long for Sent confirmation, then pruned. */
export const SENT_RETENTION_DAYS = 7
/** Newest-first outbox listing cap, per the sync spec. */
export const OUTBOX_LIST_LIMIT = 200
/** Default local-search result cap. */
export const FTS_DEFAULT_LIMIT = 50

/** Automatic outbox retry gaps after attempts 1..4; attempt 5 is terminal. */
export const OUTBOX_RETRY_DELAYS_MS = [5_000, 30_000, 120_000, 600_000] as const
export const OUTBOX_MAX_ATTEMPTS = 5

export type OutboxState = 'queued' | 'sending' | 'failed' | 'sent' | 'cancelled'
export type AttachmentState = 'staged' | 'quarantined' | 'stored'
export type IntentState = 'queued' | 'acked' | 'failed'

export interface OutboxRecord {
  id: string
  accountId: string
  req: SendRequest
  state: OutboxState
  attempts: number
  nextRetryAt?: string
  lastError?: string
  queuedAt: string
  sentAt?: string
}

export interface AccountRecord {
  accountId: string
  label: string
  kind: 'imap' | 'gmail' | 'graph'
  sessionOnly: boolean
  credRef?: string
  createdAt: string
}

export interface FolderInput {
  id: string
  name: string
  role: MailFolder['role']
  unread?: number
}

export interface AttachmentRecord {
  accountId: string
  msgUid: string
  contentId?: string
  filename: string
  mime: string
  size: number
  blobPath: string
  state: AttachmentState
}

export interface StoredIntent extends MailIntent {
  folderId: string
  attempts: number
  createdAt: string
}

export interface MessageBody {
  text: string
  htmlRef?: string
  truncated: boolean
}

export interface MessageInput {
  folderId: string
  uid: string
  msgIdHeader?: string
  threadId: string
  from: string[]
  to: string[]
  cc?: string[]
  subject?: string
  date: string
  flags?: Partial<MailFlags>
  labels?: string[]
  snippet?: string
  bodyText?: string
  /** Full HTML; stored content-addressed on disk, never inline. */
  bodyHtml?: string
  hasAttach?: boolean
  modseq?: string
}

/** Retry gap for the next automatic attempt, or null when attempts are spent. */
export function nextOutboxRetryDelay(attempts: number): number | null {
  if (attempts < 1 || attempts > OUTBOX_RETRY_DELAYS_MS.length) return null
  return OUTBOX_RETRY_DELAYS_MS[attempts - 1]
}

/**
 * Fold a provider error to one short line for outbox lastError storage.
 * Callers pass the provider message only (never subjects, addresses,
 * bodies, or secrets); this helper only folds whitespace and caps length.
 */
export function sanitizeErrorText(message: string): string {
  const folded = message
    .replace(/[\r\n]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
  return folded.length > LAST_ERROR_CHARS ? `${folded.slice(0, LAST_ERROR_CHARS)}` : folded
}

/** Profile store root when the caller does not pin one (tests pass tmp dirs). */
export function defaultMailStoreDir(): string {
  return join(homedir(), '.genoffice', 'mail')
}

function sanitizeAccountFile(accountId: string): string {
  const safe = accountId.replace(/[^A-Za-z0-9._-]/g, '_').slice(0, 128)
  return `${safe.length > 0 ? safe : 'account'}.sqlite`
}

export function accountDbPath(rootDir: string, accountId: string): string {
  return join(rootDir, sanitizeAccountFile(accountId))
}

function ensureDirs(rootDir: string): void {
  mkdirSync(join(rootDir, 'blobs'), { recursive: true })
  mkdirSync(join(rootDir, 'attachments'), { recursive: true })
  mkdirSync(join(rootDir, 'staging'), { recursive: true })
}

function blobRelPath(shaHex: string): string {
  return ['blobs', shaHex.slice(0, 2), `${shaHex}.html`].join('/')
}

function resolveInside(rootDir: string, relPath: string): string | null {
  const abs = resolve(rootDir, relPath)
  const root = resolve(rootDir)
  if (abs !== root && !abs.startsWith(root + sep)) return null
  return abs
}

function asInt(value: unknown): number {
  return typeof value === 'bigint' ? Number(value) : (value as number)
}

function asText(value: unknown): string {
  return value === null || value === undefined ? '' : String(value)
}

function parseJsonArray(value: unknown): string[] {
  if (typeof value !== 'string' || value === '') return []
  try {
    const parsed: unknown = JSON.parse(value)
    return Array.isArray(parsed)
      ? parsed.filter((item): item is string => typeof item === 'string')
      : []
  } catch {
    return []
  }
}

function toEnvelope(accountId: string, row: Record<string, unknown>): MailEnvelope {
  void accountId
  return {
    ref: {
      uid: asText(row['uid']),
      folderId: asText(row['folderId']),
      ...(row['msgId'] !== null && row['msgId'] !== undefined
        ? { msgIdHeader: asText(row['msgId']) }
        : {}),
      threadId: asText(row['threadId']),
    },
    from: parseJsonArray(row['fromJson']),
    to: parseJsonArray(row['toJson']),
    cc: parseJsonArray(row['ccJson']),
    subject: asText(row['subject']),
    date: asText(row['dateIso']),
    flags: {
      seen: asInt(row['seen']) === 1,
      flagged: asInt(row['flagged']) === 1,
      answered: asInt(row['answered']) === 1,
      draft: asInt(row['draft']) === 1,
    },
    labels: parseJsonArray(row['labelsJson']),
    snippet: asText(row['snippet']),
    hasAttach: asInt(row['hasAttach']) === 1,
  }
}

function toOutboxRecord(row: Record<string, unknown>): OutboxRecord {
  let req: SendRequest
  try {
    req = JSON.parse(asText(row['reqJson'])) as SendRequest
  } catch {
    req = { to: [], subject: '', text: '' }
  }
  const record: OutboxRecord = {
    id: asText(row['id']),
    accountId: asText(row['accountId']),
    req,
    state: asText(row['state']) as OutboxState,
    attempts: asInt(row['attempts']),
    queuedAt: asText(row['queuedAt']),
  }
  const nextRetryAt = row['nextRetryAt']
  if (typeof nextRetryAt === 'string' && nextRetryAt !== '') record.nextRetryAt = nextRetryAt
  const lastError = row['lastError']
  if (typeof lastError === 'string' && lastError !== '') record.lastError = lastError
  const sentAt = row['sentAt']
  if (typeof sentAt === 'string' && sentAt !== '') record.sentAt = sentAt
  return record
}

/** Quote one FTS5 MATCH token; user text never reaches MATCH raw. */
function quoteFtsToken(token: string): string {
  return `"${token.replace(/"/g, '""')}"`
}

function buildFtsMatch(query: string): string | null {
  const tokens = query.split(/\s+/u).filter((token) => token !== '')
  if (tokens.length === 0) return null
  // Each token becomes a prefix phrase. CJK runs form single unicode61
  // tokens, so a short CJK query only retrieves via the trailing wildcard;
  // Latin queries gain typeahead behavior. Quoting keeps user syntax inert.
  return tokens.map((token) => `${quoteFtsToken(token)}*`).join(' AND ')
}

/** Escape a LIKE pattern; user text never reaches LIKE raw. */
function escapeLikePattern(text: string): string {
  return text.replace(/[\\%_]/g, (ch) => `\\${ch}`)
}

/** Strip in-memory attachment payloads: the store keeps paths, never blobs. */
function canonicalizeRequest(req: SendRequest): SendRequest {
  const clean: SendRequest = { to: [...req.to], subject: req.subject, text: req.text }
  if (req.cc !== undefined) clean.cc = [...req.cc]
  if (req.bcc !== undefined) clean.bcc = [...req.bcc]
  if (req.html !== undefined) clean.html = req.html
  if (req.attachments !== undefined) {
    clean.attachments = req.attachments.map((att) => ({
      filename: att.filename,
      mime: att.mime,
      size: att.size,
      path: att.path,
    }))
  }
  return clean
}

export interface OpenStoreOptions {
  rootDir: string
  /** Isolated backing file for tests; skips the shared per-account layout. */
  dbPath?: string
  clock?: () => string
}

export class SqliteMailStore implements MailStore {
  private readonly db: DatabaseSync
  private readonly accountId: string
  private readonly rootDir: string
  private readonly clock: () => string
  private closed = false

  private constructor(db: DatabaseSync, accountId: string, rootDir: string, clock: () => string) {
    this.db = db
    this.accountId = accountId
    this.rootDir = rootDir
    this.clock = clock
  }

  static open(accountId: string, opts: OpenStoreOptions): SqliteMailStore {
    if (accountId === '') throw new Error('mail store requires an account id')
    ensureDirs(opts.rootDir)
    const dbPath = opts.dbPath ?? accountDbPath(opts.rootDir, accountId)
    const db = openDatabase(dbPath)
    try {
      runMigrations(db)
    } catch (err) {
      closeDatabase(db)
      throw err
    }
    return new SqliteMailStore(
      db,
      accountId,
      opts.rootDir,
      opts.clock ?? (() => new Date().toISOString()),
    )
  }

  close(): void {
    if (this.closed) return
    this.closed = true
    closeDatabase(this.db)
  }

  get schemaVersion(): number {
    return getUserVersion(this.db)
  }

  // Accounts and folders.

  upsertAccount(record: AccountRecord): void {
    this.db
      .prepare(
        `INSERT INTO accounts (accountId, label, kind, sessionOnly, credRef, createdAt)
         VALUES (?, ?, ?, ?, ?, ?)
         ON CONFLICT (accountId) DO UPDATE SET
           label = excluded.label, kind = excluded.kind,
           sessionOnly = excluded.sessionOnly, credRef = excluded.credRef`,
      )
      .run(
        record.accountId,
        record.label,
        record.kind,
        record.sessionOnly ? 1 : 0,
        record.credRef ?? null,
        record.createdAt,
      )
  }

  getAccount(accountId: string): AccountRecord | undefined {
    if (accountId !== this.accountId) return undefined
    const row = this.db.prepare('SELECT * FROM accounts WHERE accountId = ?').get(accountId) as
      Record<string, unknown> | undefined
    if (!row) return undefined
    const record: AccountRecord = {
      accountId: asText(row['accountId']),
      label: asText(row['label']),
      kind: asText(row['kind']) as AccountRecord['kind'],
      sessionOnly: asInt(row['sessionOnly']) === 1,
      createdAt: asText(row['createdAt']),
    }
    const credRef = row['credRef']
    if (typeof credRef === 'string') record.credRef = credRef
    return record
  }

  private ensureAccountRow(): void {
    this.db
      .prepare(
        `INSERT INTO accounts (accountId, label, kind, sessionOnly, createdAt)
         VALUES (?, ?, 'imap', 0, ?) ON CONFLICT (accountId) DO NOTHING`,
      )
      .run(this.accountId, this.accountId, this.clock())
  }

  upsertFolder(accountId: string, folder: FolderInput): void {
    if (accountId !== this.accountId) throw new Error('mail store is per-account')
    this.ensureAccountRow()
    this.db
      .prepare(
        `INSERT INTO folders (accountId, folderId, name, role, unread)
         VALUES (?, ?, ?, ?, ?)
         ON CONFLICT (accountId, folderId) DO UPDATE SET
           name = excluded.name, role = excluded.role, unread = excluded.unread`,
      )
      .run(accountId, folder.id, folder.name, folder.role, folder.unread ?? 0)
  }

  listFolders(accountId: string): MailFolder[] {
    if (accountId !== this.accountId) return []
    return (
      this.db
        .prepare(
          'SELECT folderId, name, role, unread FROM folders WHERE accountId = ? ORDER BY name',
        )
        .all(accountId) as Record<string, unknown>[]
    ).map((row) => ({
      id: asText(row['folderId']),
      name: asText(row['name']),
      role: asText(row['role']) as MailFolder['role'],
      unread: asInt(row['unread']),
    }))
  }

  setFolderUnread(accountId: string, folderId: string, unread: number): void {
    if (accountId !== this.accountId) throw new Error('mail store is per-account')
    this.db
      .prepare('UPDATE folders SET unread = ? WHERE accountId = ? AND folderId = ?')
      .run(unread, accountId, folderId)
  }

  deleteFolder(accountId: string, folderId: string): void {
    if (accountId !== this.accountId) throw new Error('mail store is per-account')
    this.db
      .prepare('DELETE FROM folders WHERE accountId = ? AND folderId = ?')
      .run(accountId, folderId)
  }

  private ensureFolderRow(folderId: string): void {
    this.ensureAccountRow()
    this.db
      .prepare(
        `INSERT INTO folders (accountId, folderId, name, role, unread)
         VALUES (?, ?, ?, 'custom', 0) ON CONFLICT (accountId, folderId) DO NOTHING`,
      )
      .run(this.accountId, folderId, folderId)
  }

  // Messages: MailStore-compatible envelope writes plus full-body writes.

  upsertEnvelopes(env: MailEnvelope[]): void {
    withTransaction(this.db, () => {
      for (const item of env) {
        this.ensureFolderRow(item.ref.folderId)
        this.db
          .prepare(
            `INSERT INTO messages
               (accountId, folderId, uid, msgId, threadId, fromJson, toJson, ccJson,
                subject, dateIso, seen, flagged, answered, draft, labelsJson,
                snippet, bodyText, hasAttach, updatedAt)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '', ?, ?)
             ON CONFLICT (accountId, folderId, uid) DO UPDATE SET
               msgId = excluded.msgId, threadId = excluded.threadId,
               fromJson = excluded.fromJson, toJson = excluded.toJson, ccJson = excluded.ccJson,
               subject = excluded.subject, dateIso = excluded.dateIso, seen = excluded.seen,
               flagged = excluded.flagged, answered = excluded.answered, draft = excluded.draft,
               labelsJson = excluded.labelsJson, snippet = excluded.snippet,
               hasAttach = excluded.hasAttach, updatedAt = excluded.updatedAt`,
          )
          .run(
            this.accountId,
            item.ref.folderId,
            item.ref.uid,
            item.ref.msgIdHeader ?? null,
            item.ref.threadId,
            JSON.stringify(item.from),
            JSON.stringify(item.to),
            JSON.stringify(item.cc),
            item.subject,
            item.date,
            item.flags.seen ? 1 : 0,
            item.flags.flagged ? 1 : 0,
            item.flags.answered ? 1 : 0,
            item.flags.draft ? 1 : 0,
            JSON.stringify(item.labels),
            item.snippet,
            item.hasAttach ? 1 : 0,
            this.clock(),
          )
      }
    })
  }

  /**
   * Full sync write: envelope plus inline plaintext/snippet and an optional
   * content-addressed HTML blob. Over-cap plaintext keeps the head inline
   * and sets bodyTruncated.
   */
  upsertMessageFull(input: MessageInput): void {
    const plain = input.bodyText ?? ''
    const truncated = plain.length > INLINE_BODY_CAP_BYTES
    const inlineText = truncated ? plain.slice(0, INLINE_BODY_CAP_BYTES) : plain
    const snippetSource =
      inlineText !== '' ? inlineText : input.bodyHtml !== undefined ? stripTags(input.bodyHtml) : ''
    const snippet = input.snippet ?? makeSnippet(snippetSource, SNIPPET_CHARS)
    let bodyHtmlRef: string | null = null
    if (input.bodyHtml !== undefined) {
      bodyHtmlRef = this.writeBodyHtml(input.bodyHtml)
    }
    const flags = input.flags ?? {}
    withTransaction(this.db, () => {
      this.ensureFolderRow(input.folderId)
      this.db
        .prepare(
          `INSERT INTO messages
             (accountId, folderId, uid, msgId, threadId, fromJson, toJson, ccJson,
              subject, dateIso, seen, flagged, answered, draft, labelsJson,
              snippet, bodyText, bodyHtmlRef, bodyTruncated, hasAttach, modseq, updatedAt)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
           ON CONFLICT (accountId, folderId, uid) DO UPDATE SET
             msgId = excluded.msgId, threadId = excluded.threadId,
             fromJson = excluded.fromJson, toJson = excluded.toJson, ccJson = excluded.ccJson,
             subject = excluded.subject, dateIso = excluded.dateIso, seen = excluded.seen,
             flagged = excluded.flagged, answered = excluded.answered, draft = excluded.draft,
             labelsJson = excluded.labelsJson, snippet = excluded.snippet,
             bodyText = excluded.bodyText,
             bodyHtmlRef = COALESCE(excluded.bodyHtmlRef, messages.bodyHtmlRef),
             bodyTruncated = excluded.bodyTruncated,
             hasAttach = excluded.hasAttach, modseq = excluded.modseq,
             updatedAt = excluded.updatedAt`,
        )
        .run(
          this.accountId,
          input.folderId,
          input.uid,
          input.msgIdHeader ?? null,
          input.threadId,
          JSON.stringify(input.from),
          JSON.stringify(input.to),
          JSON.stringify(input.cc ?? []),
          input.subject ?? '',
          input.date,
          flags.seen === true ? 1 : 0,
          flags.flagged === true ? 1 : 0,
          flags.answered === true ? 1 : 0,
          flags.draft === true ? 1 : 0,
          JSON.stringify(input.labels ?? []),
          snippet,
          inlineText,
          bodyHtmlRef,
          truncated ? 1 : 0,
          input.hasAttach === true ? 1 : 0,
          input.modseq ?? null,
          this.clock(),
        )
    })
  }

  private writeBodyHtml(html: string): string {
    const sha = contentAddress(new TextEncoder().encode(html))
    const rel = blobRelPath(sha)
    const abs = resolveInside(this.rootDir, rel)
    if (!abs) throw new Error('mail store blob path escaped the store root')
    mkdirSync(join(this.rootDir, 'blobs', sha.slice(0, 2)), { recursive: true })
    if (!existsSync(abs)) writeFileSync(abs, html, 'utf8')
    return rel
  }

  readBodyHtml(htmlRef: string): string | undefined {
    const abs = resolveInside(this.rootDir, htmlRef)
    if (!abs || !existsSync(abs)) return undefined
    return readFileSync(abs, 'utf8')
  }

  getEnvelope(accountId: string, folderId: string, uid: string): MailEnvelope | undefined {
    if (accountId !== this.accountId) return undefined
    const row = this.db
      .prepare('SELECT * FROM messages WHERE accountId = ? AND folderId = ? AND uid = ?')
      .get(accountId, folderId, uid) as Record<string, unknown> | undefined
    return row ? toEnvelope(accountId, row) : undefined
  }

  getMessageBody(accountId: string, folderId: string, uid: string): MessageBody | undefined {
    if (accountId !== this.accountId) return undefined
    const row = this.db
      .prepare(
        'SELECT bodyText, bodyHtmlRef, bodyTruncated FROM messages WHERE accountId = ? AND folderId = ? AND uid = ?',
      )
      .get(accountId, folderId, uid) as Record<string, unknown> | undefined
    if (!row) return undefined
    const body: MessageBody = {
      text: asText(row['bodyText']),
      truncated: asInt(row['bodyTruncated']) === 1,
    }
    const ref = row['bodyHtmlRef']
    if (typeof ref === 'string' && ref !== '') body.htmlRef = ref
    return body
  }

  getThread(threadId: string): MailEnvelope[] {
    return (
      this.db
        .prepare('SELECT * FROM messages WHERE accountId = ? AND threadId = ? ORDER BY dateIso')
        .all(this.accountId, threadId) as Record<string, unknown>[]
    ).map((row) => toEnvelope(this.accountId, row))
  }

  listEnvelopes(
    accountId: string,
    folderId: string,
    opts?: { limit?: number; offset?: number },
  ): MailEnvelope[] {
    if (accountId !== this.accountId) return []
    const limit = opts?.limit ?? 100
    const offset = opts?.offset ?? 0
    return (
      this.db
        .prepare(
          `SELECT * FROM messages WHERE accountId = ? AND folderId = ?
           ORDER BY dateIso DESC LIMIT ? OFFSET ?`,
        )
        .all(accountId, folderId, limit, offset) as Record<string, unknown>[]
    ).map((row) => toEnvelope(this.accountId, row))
  }

  setMessageFlags(
    accountId: string,
    folderId: string,
    uid: string,
    flags: Partial<MailFlags>,
  ): void {
    if (accountId !== this.accountId) throw new Error('mail store is per-account')
    const sets: string[] = []
    const values: Array<string | number> = []
    if (flags.seen !== undefined) {
      sets.push('seen = ?')
      values.push(flags.seen ? 1 : 0)
    }
    if (flags.flagged !== undefined) {
      sets.push('flagged = ?')
      values.push(flags.flagged ? 1 : 0)
    }
    if (flags.answered !== undefined) {
      sets.push('answered = ?')
      values.push(flags.answered ? 1 : 0)
    }
    if (flags.draft !== undefined) {
      sets.push('draft = ?')
      values.push(flags.draft ? 1 : 0)
    }
    if (sets.length === 0) return
    sets.push('updatedAt = ?')
    values.push(this.clock())
    this.db
      .prepare(
        `UPDATE messages SET ${sets.join(', ')} WHERE accountId = ? AND folderId = ? AND uid = ?`,
      )
      .run(...values, accountId, folderId, uid)
  }

  setMessageLabels(accountId: string, folderId: string, uid: string, labels: string[]): void {
    if (accountId !== this.accountId) throw new Error('mail store is per-account')
    this.db
      .prepare(
        'UPDATE messages SET labelsJson = ?, updatedAt = ? WHERE accountId = ? AND folderId = ? AND uid = ?',
      )
      .run(JSON.stringify(labels), this.clock(), accountId, folderId, uid)
  }

  moveMessage(accountId: string, ref: MailMessageRef, toFolderId: string): void {
    if (accountId !== this.accountId) throw new Error('mail store is per-account')
    withTransaction(this.db, () => {
      this.ensureFolderRow(toFolderId)
      this.db
        .prepare(
          `UPDATE messages SET folderId = ?, updatedAt = ?
           WHERE accountId = ? AND folderId = ? AND uid = ?`,
        )
        .run(toFolderId, this.clock(), accountId, ref.folderId, ref.uid)
    })
  }

  setMessageModseq(accountId: string, folderId: string, uid: string, modseq: string): void {
    if (accountId !== this.accountId) throw new Error('mail store is per-account')
    this.db
      .prepare('UPDATE messages SET modseq = ? WHERE accountId = ? AND folderId = ? AND uid = ?')
      .run(modseq, accountId, folderId, uid)
  }

  deleteMessage(accountId: string, folderId: string, uid: string): void {
    if (accountId !== this.accountId) throw new Error('mail store is per-account')
    this.db
      .prepare('DELETE FROM messages WHERE accountId = ? AND folderId = ? AND uid = ?')
      .run(accountId, folderId, uid)
  }

  /**
   * Folder purge for full resync (UIDVALIDITY change, expired history/delta).
   * Deletes message rows only; the FTS triggers mirror the delete, while
   * outbox, drafts, and intents are untouched by construction.
   */
  purgeFolderMessages(accountId: string, folderId: string): number {
    if (accountId !== this.accountId) throw new Error('mail store is per-account')
    return withTransaction(this.db, () => {
      const result = this.db
        .prepare('DELETE FROM messages WHERE accountId = ? AND folderId = ?')
        .run(accountId, folderId)
      return Number(result.changes)
    })
  }

  /**
   * SyncStorePort.purgeFolder alias (T+Y contract, BUG-1): the sync engine
   * calls purgeFolder on full resync. Row DELETEs are mirrored into the
   * FTS index by the messages_ad trigger, so no rebuildFts is needed here.
   */
  purgeFolder(accountId: string, folderId: string): void {
    this.purgeFolderMessages(accountId, folderId)
  }

  /**
   * Rebuild the FTS mirror from the messages table (runs inside resync).
   * Manual delete plus reinsert: the FTS column `fromText` maps from the
   * content column `fromJson`, so the built-in `rebuild` command (which
   * selects same-named columns) cannot run here.
   */
  rebuildFts(): void {
    withTransaction(this.db, () => {
      this.db.exec(`INSERT INTO messages_fts(messages_fts) VALUES('delete-all')`)
      this.db.exec(
        `INSERT INTO messages_fts(rowid, subject, fromText, snippet, bodyText)
         SELECT rowid, subject, fromJson, snippet, bodyText FROM messages`,
      )
    })
  }

  fts(query: string, accountIds: string[]): MailEnvelope[] {
    if (!accountIds.includes(this.accountId)) return []
    return this.searchLocal(query, { limit: FTS_DEFAULT_LIMIT })
  }

  searchLocal(query: string, opts?: { limit?: number }): MailEnvelope[] {
    const match = buildFtsMatch(query)
    if (!match) return []
    const limit = opts?.limit ?? FTS_DEFAULT_LIMIT
    const ranked = (
      this.db
        .prepare(
          `SELECT m.* FROM messages_fts JOIN messages m ON m.rowid = messages_fts.rowid
           WHERE m.accountId = ? AND messages_fts MATCH ? ORDER BY bm25(messages_fts) LIMIT ?`,
        )
        .all(this.accountId, match, limit) as Record<string, unknown>[]
    ).map((row) => toEnvelope(this.accountId, row))
    if (ranked.length > 0) return ranked
    // Substring fallback: unicode61 groups CJK runs into single tokens, so
    // a mid-string CJK (or mid-token Latin) query cannot match by prefix.
    // Only reached when ranked search is empty; LIKE patterns stay escaped.
    const like = `%${escapeLikePattern(query.trim())}%`
    return (
      this.db
        .prepare(
          `SELECT * FROM messages WHERE accountId = ?
           AND (subject LIKE ? ESCAPE '\\' OR fromJson LIKE ? ESCAPE '\\'
             OR snippet LIKE ? ESCAPE '\\' OR bodyText LIKE ? ESCAPE '\\')
           ORDER BY dateIso DESC LIMIT ?`,
        )
        .all(this.accountId, like, like, like, like, limit) as Record<string, unknown>[]
    ).map((row) => toEnvelope(this.accountId, row))
  }

  // Attachments: state machine only; blobs live on disk.

  addAttachment(record: AttachmentRecord): void {
    if (record.accountId !== this.accountId) throw new Error('mail store is per-account')
    this.db
      .prepare(
        `INSERT INTO attachments (accountId, msgUid, contentId, filename, mime, size, blobPath, state)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)
         ON CONFLICT (accountId, msgUid, filename) DO UPDATE SET
           contentId = excluded.contentId, mime = excluded.mime, size = excluded.size,
           blobPath = excluded.blobPath, state = excluded.state`,
      )
      .run(
        record.accountId,
        record.msgUid,
        record.contentId ?? null,
        record.filename,
        record.mime,
        record.size,
        record.blobPath,
        record.state,
      )
  }

  listAttachments(accountId: string, msgUid: string): AttachmentRecord[] {
    if (accountId !== this.accountId) return []
    return (
      this.db
        .prepare('SELECT * FROM attachments WHERE accountId = ? AND msgUid = ? ORDER BY filename')
        .all(accountId, msgUid) as Record<string, unknown>[]
    ).map((row) => {
      const record: AttachmentRecord = {
        accountId: asText(row['accountId']),
        msgUid: asText(row['msgUid']),
        filename: asText(row['filename']),
        mime: asText(row['mime']),
        size: asInt(row['size']),
        blobPath: asText(row['blobPath']),
        state: asText(row['state']) as AttachmentState,
      }
      const contentId = row['contentId']
      if (typeof contentId === 'string') record.contentId = contentId
      return record
    })
  }

  /**
   * Record a quarantine transition. Promoting to stored requires a relative
   * in-store path; staged rows may point at absolute temp staging paths.
   */
  setAttachmentState(
    accountId: string,
    msgUid: string,
    filename: string,
    state: AttachmentState,
    blobPath?: string,
  ): void {
    if (accountId !== this.accountId) throw new Error('mail store is per-account')
    if (state === 'stored') {
      const next = blobPath ?? this.getAttachmentPath(accountId, msgUid, filename)
      if (!next || next === '' || next.startsWith('/') || next.includes('..')) {
        throw new Error('stored attachments require a relative in-store path')
      }
      this.db
        .prepare(
          'UPDATE attachments SET state = ?, blobPath = ? WHERE accountId = ? AND msgUid = ? AND filename = ?',
        )
        .run(state, next, accountId, msgUid, filename)
      return
    }
    if (blobPath !== undefined) {
      this.db
        .prepare(
          'UPDATE attachments SET state = ?, blobPath = ? WHERE accountId = ? AND msgUid = ? AND filename = ?',
        )
        .run(state, blobPath, accountId, msgUid, filename)
    } else {
      this.db
        .prepare(
          'UPDATE attachments SET state = ? WHERE accountId = ? AND msgUid = ? AND filename = ?',
        )
        .run(state, accountId, msgUid, filename)
    }
  }

  private getAttachmentPath(
    accountId: string,
    msgUid: string,
    filename: string,
  ): string | undefined {
    const row = this.db
      .prepare(
        'SELECT blobPath FROM attachments WHERE accountId = ? AND msgUid = ? AND filename = ?',
      )
      .get(accountId, msgUid, filename) as Record<string, unknown> | undefined
    const path = row?.['blobPath']
    return typeof path === 'string' ? path : undefined
  }

  deleteAttachment(accountId: string, msgUid: string, filename: string): void {
    if (accountId !== this.accountId) throw new Error('mail store is per-account')
    this.db
      .prepare('DELETE FROM attachments WHERE accountId = ? AND msgUid = ? AND filename = ?')
      .run(accountId, msgUid, filename)
  }

  // Sync state.

  setSyncState(state: SyncState): void {
    if (state.accountId !== this.accountId) throw new Error('mail store is per-account')
    this.db
      .prepare(
        `INSERT INTO sync_state
           (accountId, folderId, uidValidity, highestModseq, historyId, deltaToken, lastSyncAt)
         VALUES (?, ?, ?, ?, ?, ?, ?)
         ON CONFLICT (accountId, folderId) DO UPDATE SET
           uidValidity = excluded.uidValidity, highestModseq = excluded.highestModseq,
           historyId = excluded.historyId, deltaToken = excluded.deltaToken,
           lastSyncAt = excluded.lastSyncAt`,
      )
      .run(
        state.accountId,
        state.folderId,
        state.uidValidity ?? null,
        state.highestModseq ?? null,
        state.historyId ?? null,
        state.deltaToken ?? null,
        state.lastSyncAt,
      )
  }

  getSyncState(accountId: string, folderId: string): SyncState | undefined {
    if (accountId !== this.accountId) return undefined
    const row = this.db
      .prepare('SELECT * FROM sync_state WHERE accountId = ? AND folderId = ?')
      .get(accountId, folderId) as Record<string, unknown> | undefined
    if (!row) return undefined
    const state: SyncState = {
      accountId: asText(row['accountId']),
      folderId: asText(row['folderId']),
      lastSyncAt: asText(row['lastSyncAt']),
    }
    const uidValidity = row['uidValidity']
    if (typeof uidValidity === 'number' || typeof uidValidity === 'bigint')
      state.uidValidity = Number(uidValidity)
    for (const key of ['highestModseq', 'historyId', 'deltaToken'] as const) {
      const value = row[key]
      if (typeof value === 'string') state[key] = value
    }
    return state
  }

  // Outbox: FIFO per account, one in-flight send, idempotent ids.

  enqueueSend(req: SendRequest): string {
    return this.enqueueOutbox(this.accountId, req)
  }

  enqueueOutbox(accountId: string, req: SendRequest, opts?: { id?: string }): string {
    if (accountId !== this.accountId) throw new Error('mail store is per-account')
    this.ensureAccountRow()
    const id = opts?.id ?? `outbox-${randomUUID()}`
    this.db
      .prepare(
        `INSERT INTO outbox (id, accountId, reqJson, state, attempts, queuedAt)
         VALUES (?, ?, ?, 'queued', 0, ?)`,
      )
      .run(id, accountId, JSON.stringify(canonicalizeRequest(req)), this.clock())
    return id
  }

  getOutbox(id: string): OutboxRecord | undefined {
    const row = this.db
      .prepare('SELECT * FROM outbox WHERE id = ? AND accountId = ?')
      .get(id, this.accountId) as Record<string, unknown> | undefined
    return row ? toOutboxRecord(row) : undefined
  }

  listOutbox(accountId: string, opts?: { limit?: number }): OutboxRecord[] {
    if (accountId !== this.accountId) return []
    const limit = Math.min(opts?.limit ?? OUTBOX_LIST_LIMIT, OUTBOX_LIST_LIMIT)
    return (
      this.db
        .prepare('SELECT * FROM outbox WHERE accountId = ? ORDER BY queuedAt DESC, id DESC LIMIT ?')
        .all(accountId, limit) as Record<string, unknown>[]
    ).map(toOutboxRecord)
  }

  outboxCounts(accountId: string): { queued: number; failed: number } {
    if (accountId !== this.accountId) return { queued: 0, failed: 0 }
    const rows = this.db
      .prepare(
        `SELECT state, COUNT(*) AS n FROM outbox WHERE accountId = ? AND state IN ('queued', 'failed') GROUP BY state`,
      )
      .all(accountId) as Record<string, unknown>[]
    let queued = 0
    let failed = 0
    for (const row of rows) {
      if (asText(row['state']) === 'queued') queued = asInt(row['n'])
      if (asText(row['state']) === 'failed') failed = asInt(row['n'])
    }
    return { queued, failed }
  }

  /**
   * Claim the oldest queued send for one in-flight transmission per account.
   * Marks the row sending and bumps attempts inside one transaction.
   */
  claimNextOutbox(accountId: string): OutboxRecord | undefined {
    if (accountId !== this.accountId) throw new Error('mail store is per-account')
    return withTransaction(this.db, () => {
      const row = this.db
        .prepare(
          `SELECT * FROM outbox WHERE accountId = ? AND state = 'queued'
           ORDER BY queuedAt, id LIMIT 1`,
        )
        .get(accountId) as Record<string, unknown> | undefined
      if (!row) return undefined
      this.db
        .prepare(`UPDATE outbox SET state = 'sending', attempts = attempts + 1 WHERE id = ?`)
        .run(asText(row['id']))
      const claimed = this.db
        .prepare('SELECT * FROM outbox WHERE id = ?')
        .get(asText(row['id'])) as Record<string, unknown> | undefined
      return claimed ? toOutboxRecord(claimed) : undefined
    })
  }

  markOutboxSent(id: string, sentAt?: string): void {
    this.db
      .prepare(
        `UPDATE outbox SET state = 'sent', sentAt = ?, nextRetryAt = NULL WHERE id = ? AND accountId = ?`,
      )
      .run(sentAt ?? this.clock(), id, this.accountId)
  }

  failOutbox(id: string, lastError: string, nextRetryAt?: string): void {
    this.db
      .prepare(
        `UPDATE outbox SET state = 'failed', lastError = ?, nextRetryAt = ? WHERE id = ? AND accountId = ?`,
      )
      .run(sanitizeErrorText(lastError), nextRetryAt ?? null, id, this.accountId)
  }

  /** Manual retry: reset attempts and re-queue at head. */
  retryOutbox(id: string): void {
    this.db
      .prepare(
        `UPDATE outbox SET state = 'queued', attempts = 0, nextRetryAt = NULL, lastError = NULL
         WHERE id = ? AND accountId = ? AND state IN ('queued', 'failed', 'cancelled')`,
      )
      .run(id, this.accountId)
  }

  /** Cancel a queued or failed entry; false when already sent or sending. */
  cancelOutbox(id: string): boolean {
    const result = this.db
      .prepare(
        `UPDATE outbox SET state = 'cancelled', nextRetryAt = NULL
         WHERE id = ? AND accountId = ? AND state IN ('queued', 'failed')`,
      )
      .run(id, this.accountId)
    return Number(result.changes) === 1
  }

  deleteOutbox(id: string): void {
    this.db.prepare('DELETE FROM outbox WHERE id = ? AND accountId = ?').run(id, this.accountId)
  }

  pruneSentOutbox(accountId: string, beforeIso: string): number {
    if (accountId !== this.accountId) throw new Error('mail store is per-account')
    const result = this.db
      .prepare(`DELETE FROM outbox WHERE accountId = ? AND state = 'sent' AND sentAt < ?`)
      .run(accountId, beforeIso)
    return Number(result.changes)
  }

  // Drafts.

  saveDraft(accountId: string, threadKey: string, text: string): void {
    if (accountId !== this.accountId) throw new Error('mail store is per-account')
    this.db
      .prepare(
        `INSERT INTO drafts (accountId, threadKey, text, updatedAt)
         VALUES (?, ?, ?, ?)
         ON CONFLICT (accountId, threadKey) DO UPDATE SET text = excluded.text, updatedAt = excluded.updatedAt`,
      )
      .run(accountId, threadKey, text, this.clock())
  }

  draftFor(accountId: string, threadKey: string): string | undefined {
    if (accountId !== this.accountId) return undefined
    const row = this.db
      .prepare('SELECT text FROM drafts WHERE accountId = ? AND threadKey = ?')
      .get(accountId, threadKey) as Record<string, unknown> | undefined
    const text = row?.['text']
    return typeof text === 'string' ? text : undefined
  }

  deleteDraft(accountId: string, threadKey: string): void {
    if (accountId !== this.accountId) throw new Error('mail store is per-account')
    this.db
      .prepare('DELETE FROM drafts WHERE accountId = ? AND threadKey = ?')
      .run(accountId, threadKey)
  }

  // Offline intents: FIFO replay queue with idempotent keys (see sync.ts).

  queueIntent(intent: Omit<MailIntent, 'id' | 'state'> & { folderId?: string }): string {
    if (intent.accountId !== this.accountId) throw new Error('mail store is per-account')
    const id = `intent-${randomUUID()}`
    this.db
      .prepare(
        `INSERT INTO intents
           (id, accountId, kind, targetUid, folderId, expectedModseq, payloadJson, state, attempts, createdAt)
         VALUES (?, ?, ?, ?, ?, ?, ?, 'queued', 0, ?)`,
      )
      .run(
        id,
        intent.accountId,
        intent.kind,
        intent.targetUid,
        intent.folderId ?? '',
        intent.expectedModseq ?? null,
        intent.payloadJson,
        this.clock(),
      )
    return id
  }

  pendingIntents(accountId: string): MailIntent[] {
    return this.listIntents(accountId, { state: 'queued' })
  }

  listIntents(accountId: string, opts?: { state?: IntentState }): StoredIntent[] {
    if (accountId !== this.accountId) return []
    const sql =
      opts?.state !== undefined
        ? 'SELECT * FROM intents WHERE accountId = ? AND state = ? ORDER BY createdAt, id'
        : 'SELECT * FROM intents WHERE accountId = ? ORDER BY createdAt, id'
    const args = opts?.state !== undefined ? [accountId, opts.state] : [accountId]
    return (this.db.prepare(sql).all(...args) as Record<string, unknown>[]).map((row) => ({
      id: asText(row['id']),
      accountId: asText(row['accountId']),
      kind: asText(row['kind']) as StoredIntent['kind'],
      targetUid: asText(row['targetUid']),
      folderId: asText(row['folderId']),
      expectedModseq:
        typeof row['expectedModseq'] === 'string' ? (row['expectedModseq'] as string) : undefined,
      payloadJson: asText(row['payloadJson']),
      state: asText(row['state']) as IntentState,
      attempts: asInt(row['attempts']),
      createdAt: asText(row['createdAt']),
    }))
  }

  markIntent(id: string, state: 'acked' | 'failed'): void {
    this.db
      .prepare(`UPDATE intents SET state = ? WHERE id = ? AND accountId = ?`)
      .run(state, id, this.accountId)
  }

  bumpIntentAttempt(id: string): void {
    this.db
      .prepare(`UPDATE intents SET attempts = attempts + 1 WHERE id = ? AND accountId = ?`)
      .run(id, this.accountId)
  }

  deleteIntent(id: string): void {
    this.db.prepare('DELETE FROM intents WHERE id = ? AND accountId = ?').run(id, this.accountId)
  }
}

const openStores = new Map<string, SqliteMailStore>()

/** Open (or reuse) the SQLite store for one account. Y lane owns the lifecycle. */
export function openMailStore(accountId: string, opts: OpenStoreOptions): SqliteMailStore {
  const key = `${opts.dbPath ?? accountDbPath(opts.rootDir, accountId)}`
  const existing = openStores.get(key)
  if (existing) return existing
  const store = SqliteMailStore.open(accountId, opts)
  openStores.set(key, store)
  return store
}

export function closeMailStore(
  accountId: string,
  opts: Pick<OpenStoreOptions, 'rootDir'> & { dbPath?: string },
): void {
  const key = `${opts.dbPath ?? accountDbPath(opts.rootDir, accountId)}`
  const existing = openStores.get(key)
  if (existing) {
    openStores.delete(key)
    existing.close()
  }
}

export function closeAllMailStores(): void {
  for (const store of openStores.values()) store.close()
  openStores.clear()
}

/**
 * Account removal: close the cached handle and delete the per-account
 * database file (one rm per the spec layout). Shared content-addressed
 * blob/attachment dirs are left alone; orphan blobs are inert.
 */
export function deleteAccountStore(
  accountId: string,
  opts: Pick<OpenStoreOptions, 'rootDir'> & { dbPath?: string },
): void {
  closeMailStore(accountId, opts)
  const file = opts.dbPath ?? accountDbPath(opts.rootDir, accountId)
  if (existsSync(file)) rmSync(file)
}
