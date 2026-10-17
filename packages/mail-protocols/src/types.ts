/**
 * Mail protocol interfaces: MailProvider (spec section 3.2) and MailStore
 * (spec section 4.2). Shapes follow the amended arch spec verbatim; the few
 * derived types below (MailQuery, MailEvent, MailIntent, SyncState,
 * StagedAttachment) carry only the fields the spec text implies. Drivers
 * implement MailProvider as pure-TS classes with injected socket/auth.
 * Secrets-injection and scheduling stay in apps/mail/src/main.
 */

export interface MailFolder {
  id: string
  name: string
  role: 'inbox' | 'sent' | 'drafts' | 'trash' | 'spam' | 'archive' | 'custom'
  unread: number
}

export interface MailMessageRef {
  uid: string
  folderId: string
  msgIdHeader?: string
  threadId: string
  /**
   * Optional exact-match hint (H1 D-H1-C3, H4 implements): the scope
   * account that produced the citation this ref was cited from. The
   * retrieval core resolves exactly there and fails loud on stale hints;
   * absent means deterministic first-scope-match fallback.
   */
  accountId?: string
}

export interface MailFlags {
  seen: boolean
  flagged: boolean
  answered: boolean
  draft: boolean
}

export interface MailEnvelope {
  ref: MailMessageRef
  from: string[]
  to: string[]
  cc: string[]
  subject: string
  date: string
  flags: MailFlags
  labels: string[]
  snippet: string
  hasAttach: boolean
}

/** Full HTML lives on disk, content-addressed; never in the DB. */
export interface MailBody {
  text: string
  htmlRef?: string
}

export interface StagedAttachment {
  filename: string
  mime: string
  size: number
  /**
   * Content-addressed store path (temp staging until quarantine scan).
   * Canonical payload: stageAttachment always populates this, and every
   * driver streams from disk at send time. Never a full blob in SQLite.
   */
  path: string
  /**
   * Memory payload for tests and tiny inline parts. Fallback only: every
   * driver prefers path when both are present, and production sends always
   * stage to disk first (data never bypasses the quarantine scan).
   */
  data?: Uint8Array
}

export interface SendRequest {
  to: string[]
  cc?: string[]
  bcc?: string[]
  subject: string
  text: string
  html?: string
  attachments?: StagedAttachment[]
}

/**
 * Local-first search; server expansion is always an explicit driver call.
 *
 * Canonical P1 shape (all fields optional). P3 sync-engine contract:
 * text/from/since/onlyUnread run server-side on every driver; the other
 * six run server-side on IMAP (SEARCH builder in imap-transport.ts) and as
 * the local matchesQuery post-filter (src/query.ts) elsewhere, so every
 * field is honored on every driver one way or the other.
 */
export interface MailQuery {
  text?: string
  from?: string
  to?: string
  subject?: string
  seen?: boolean
  flagged?: boolean
  hasAttach?: boolean
  since?: string
  before?: string
  onlyUnread?: boolean
}

/**
 * Push events delivered over mail:sync-events (new mail, flag changes,
 * sync state). Payloads carry refs and flags only, never bodies or secrets.
 *
 * Canonical P1 shape: single `kind` discriminant (7 variants). The IMAP
 * driver's former `type`-discriminated names map onto it without loss:
 * message-new -> messages-added (single ref), message-updated ->
 * flags-changed (flags filled when the poll row carries them, else {}),
 * message-deleted -> message-deleted, resync -> resync,
 * conflict -> conflict, sync-error -> sync-state with state 'error'.
 */
export type MailEvent =
  | { kind: 'messages-added'; accountId: string; folderId: string; refs: MailMessageRef[] }
  | { kind: 'flags-changed'; accountId: string; ref: MailMessageRef; flags: Partial<MailFlags> }
  | { kind: 'message-moved'; accountId: string; ref: MailMessageRef; toFolderId: string }
  | { kind: 'message-deleted'; accountId: string; ref: MailMessageRef }
  | {
      kind: 'sync-state'
      accountId: string
      folderId: string
      state: 'syncing' | 'idle' | 'error'
      detail?: string
    }
  | {
      kind: 'resync'
      accountId: string
      folderId: string
      reason: 'uidvalidity-changed' | 'modseq-expired' | 'manual'
    }
  | { kind: 'conflict'; accountId: string; folderId: string; uid: string; detail: string }

export interface MailProvider {
  connect(): Promise<void>
  disconnect(): Promise<void>
  listFolders(): Promise<MailFolder[]>
  list(folderId: string, range: { afterUid?: string; limit: number }): Promise<MailEnvelope[]>
  search(folderId: string, query: MailQuery): Promise<MailEnvelope[]>
  fetch(ref: MailMessageRef): Promise<MailBody>
  /** Enqueue into the outbox; the sender transmits. Returns the outbox id. */
  send(req: SendRequest): Promise<{ outboxId: string }>
  move(ref: MailMessageRef, toFolderId: string): Promise<void>
  setFlags(ref: MailMessageRef, flags: Partial<MailFlags>): Promise<void>
  delete(ref: MailMessageRef): Promise<void>
  /** IDLE / historyId / delta mechanics live inside drivers. */
  watch(handler: (ev: MailEvent) => void): { stop(): void }
}

/** Offline intent queued while disconnected; replayed idempotently. */
export interface MailIntent {
  id: string
  accountId: string
  kind: 'move' | 'flag' | 'delete' | 'label'
  targetUid: string
  /** Stale modseq means re-fetch then retry once, else surface conflict. */
  expectedModseq?: string
  payloadJson: string
  state: 'queued' | 'acked' | 'failed'
}

export interface SyncState {
  accountId: string
  folderId: string
  uidValidity?: number
  highestModseq?: string
  historyId?: string
  deltaToken?: string
  lastSyncAt: string
}

/**
 * Replayable local store (spec section 4.2). Plaintext plus snippet stay
 * inline; full HTML is content-addressed on disk; attachments stream to
 * disk (staged temp, then quarantine scan, then store), never full blobs
 * in SQLite. credRef is an opaque vault handle, never a secret.
 */
export interface MailStore {
  upsertEnvelopes(env: MailEnvelope[]): void
  getThread(threadId: string): MailEnvelope[]
  /** Outbox id for the enqueued send. */
  enqueueSend(req: SendRequest): string
  saveDraft(accountId: string, threadKey: string, text: string): void
  queueIntent(it: Omit<MailIntent, 'id' | 'state'>): string
  pendingIntents(accountId: string): MailIntent[]
  markIntent(id: string, state: 'acked' | 'failed'): void
  setSyncState(s: SyncState): void
  getSyncState(accountId: string, folderId: string): SyncState | undefined
  /** Local FTS first; server expansion stays explicit. */
  fts(query: string, accountIds: string[]): MailEnvelope[]
}
