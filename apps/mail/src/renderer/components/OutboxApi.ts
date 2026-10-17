import type { MailEvent, SendRequest } from '@genoffice/mail-protocols'

/**
 * Outbox + offline-status renderer contract (Lane O owns this file).
 * Mirrors the P4 sync spec sections 5.1 (OutboxRecord, five states) and 6
 * (offline-status subscribe) plus Lane Y's MailApi additions
 * (`syncNow`/`outboxList`/`outboxRetry`/`outboxCancel`/`onOfflineStatus`).
 * The shapes are restated here so this lane compiles against the P3
 * baseline (Y appends the canonical names to `shared/ipc.ts` on its own
 * branch); `resolveOutboxSyncApi` binds to the live bridge when present.
 * No secrets or message bodies cross this surface — refs, counts, and
 * main-side sanitized error text only.
 */

/** The five outbox states the UI may rely on (spec section 5.1). */
export type OutboxState = 'queued' | 'sending' | 'failed' | 'sent' | 'cancelled'

/** Outbox entry as the renderer sees it (spec section 5.1). */
export interface OutboxRecord {
  id: string
  accountId: string
  req: SendRequest
  state: OutboxState
  attempts: number
  nextRetryAt?: string
  /** Sanitized main-side; the renderer displays it verbatim, never logs it. */
  lastError?: string
  queuedAt: string
  sentAt?: string
}

/** Offline status event payload (spec section 6). */
export interface OfflineStatus {
  online: boolean
  accountId?: string
  detail?: string
}

/**
 * The P4 sync/outbox slice of the preload bridge. Components take this
 * interface (not the global) so tests can inject `FakeOutboxSyncApi`.
 */
export interface OutboxSyncApi {
  /** Trigger an incremental sync step; false when one already runs or offline. */
  syncNow(accountId: string, folderId?: string): Promise<{ started: boolean }>
  /** Outbox entries newest-first (capped at 200). */
  outboxList(accountId: string): Promise<OutboxRecord[]>
  /** Manual retry: reset the budget and re-queue at head. */
  outboxRetry(accountId: string, outboxId: string): Promise<{ outboxId: string; state: string }>
  /** Cancel a queued or failed entry; false when already sent/sending. */
  outboxCancel(accountId: string, outboxId: string): Promise<{ cancelled: boolean }>
  /** Offline status; the bridge replays the current value on subscribe. */
  onOfflineStatus(handler: (status: OfflineStatus) => void): () => void
  /** Existing sync-event fan-out (refs/flags only, never bodies or secrets). */
  onSyncEvent(handler: (event: MailEvent) => void): () => void
}

type BridgeCandidate = Partial<OutboxSyncApi> | undefined

function isOutboxSyncApi(candidate: BridgeCandidate): candidate is OutboxSyncApi {
  if (!candidate || typeof candidate !== 'object') return false
  return (
    typeof candidate.syncNow === 'function' &&
    typeof candidate.outboxList === 'function' &&
    typeof candidate.outboxRetry === 'function' &&
    typeof candidate.outboxCancel === 'function' &&
    typeof candidate.onOfflineStatus === 'function' &&
    typeof candidate.onSyncEvent === 'function'
  )
}

/**
 * Bind to the live preload bridge when Lane Y's methods exist; otherwise
 * null (P3 bridge or plain vite dev). Callers fall back to a fake or hide
 * the P4 affordances — never throw from render.
 */
export function resolveOutboxSyncApi(): OutboxSyncApi | null {
  if (typeof window === 'undefined') return null
  const candidate = (window as unknown as { mailApi?: BridgeCandidate }).mailApi
  return isOutboxSyncApi(candidate) ? candidate : null
}

/** Seed for one fake outbox entry (example.com addresses only). */
export interface FakeOutboxSeed {
  id: string
  subject: string
  to: string[]
  state: OutboxState
  attempts?: number
  lastError?: string
  queuedAt?: string
}

function seedToRecord(accountId: string, index: number, seed: FakeOutboxSeed): OutboxRecord {
  return {
    id: seed.id,
    accountId,
    req: { to: seed.to, subject: seed.subject, text: '' },
    state: seed.state,
    attempts: seed.attempts ?? 0,
    lastError: seed.lastError,
    queuedAt: seed.queuedAt ?? `2026-10-15T04:00:${String(index).padStart(2, '0')}Z`,
  }
}

/**
 * In-memory OutboxSyncApi for component tests and vite dev. Holds per-test
 * records plus scripted offline/sync events; drives nothing real.
 */
export class FakeOutboxSyncApi implements OutboxSyncApi {
  private readonly accountId: string
  private records: OutboxRecord[]
  private online = true
  private offlineDetail?: string
  private readonly offlineHandlers = new Set<(status: OfflineStatus) => void>()
  private readonly syncHandlers = new Set<(event: MailEvent) => void>()
  readonly syncNowCalls: Array<{ accountId: string; folderId?: string }> = []

  constructor(accountId = 'demo', seeds: FakeOutboxSeed[] = []) {
    this.accountId = accountId
    this.records = seeds.map((seed, index) => seedToRecord(accountId, index, seed))
  }

  async syncNow(accountId: string, folderId?: string): Promise<{ started: boolean }> {
    this.syncNowCalls.push({ accountId, folderId })
    return { started: this.online }
  }

  async outboxList(accountId: string): Promise<OutboxRecord[]> {
    void accountId
    return [...this.records].sort((a, b) =>
      a.queuedAt === b.queuedAt ? (a.id < b.id ? 1 : -1) : a.queuedAt < b.queuedAt ? 1 : -1,
    )
  }

  async outboxRetry(
    _accountId: string,
    outboxId: string,
  ): Promise<{ outboxId: string; state: string }> {
    const record = this.records.find((item) => item.id === outboxId)
    if (record && (record.state === 'failed' || record.state === 'cancelled')) {
      record.state = 'queued'
      record.attempts = 0
      record.lastError = undefined
      this.emitSync({
        kind: 'sync-state',
        accountId: this.accountId,
        folderId: 'outbox',
        state: 'syncing',
      })
    }
    return { outboxId, state: 'queued' }
  }

  async outboxCancel(_accountId: string, outboxId: string): Promise<{ cancelled: boolean }> {
    const record = this.records.find((item) => item.id === outboxId)
    if (record && (record.state === 'queued' || record.state === 'failed')) {
      record.state = 'cancelled'
      return { cancelled: true }
    }
    return { cancelled: false }
  }

  onOfflineStatus(handler: (status: OfflineStatus) => void): () => void {
    this.offlineHandlers.add(handler)
    // Replay the current value on subscribe (spec section 6).
    handler({ online: this.online, accountId: this.accountId, detail: this.offlineDetail })
    return () => {
      this.offlineHandlers.delete(handler)
    }
  }

  onSyncEvent(handler: (event: MailEvent) => void): () => void {
    this.syncHandlers.add(handler)
    return () => {
      this.syncHandlers.delete(handler)
    }
  }

  /** Flip connectivity and notify subscribers (test/dev driver). */
  setOnline(online: boolean, detail?: string): void {
    this.online = online
    this.offlineDetail = detail
    const status: OfflineStatus = { online, accountId: this.accountId, detail }
    for (const handler of this.offlineHandlers) handler(status)
  }

  /** Emit a sync event to subscribers (test/dev driver). */
  emitSync(event: MailEvent): void {
    for (const handler of this.syncHandlers) handler(event)
  }

  /** Queue a record (test/dev driver). */
  addRecord(seed: FakeOutboxSeed): void {
    this.records.push(seedToRecord(this.accountId, this.records.length, seed))
  }
}
