import {
  decideSync,
  type IntentAttempt,
  type MailEnvelope,
  type MailEvent,
  type MailIntent,
  type MailStore,
  type SendRequest,
  type SyncDecision,
  type SyncSignals,
  type SyncState,
} from '@genoffice/mail-protocols'
import { checkOutboxCaps, type AccountKind } from './caps'
import { replayIntents } from './conflicts'
import { OutboxCapError, OutboxQueue, type OutboxRecord, type OutboxStatus } from './outbox'
import { nextPollDelay, POLL_BASE_MS, POLL_CEILING_MS } from './poll'
import { sanitizeError } from './redact'

/**
 * Incremental sync + outbox replay engine (P4 spec sections 3-5). Runs in
 * the main process only; the renderer talks to it through the
 * typed+zod-validated sync/outbox IPC channels (see ./ipc.ts).
 *
 * The engine is driver-agnostic behind SyncDriverPort: IMAP (IDLE where
 * advertised else 30 s poll), Gmail (history.list from stored historyId,
 * expired id means full resync), and Graph (delta links, expiry means
 * full re-list) all plug in through the same seam. Tests inject scripted
 * fakes; no real accounts or credentials appear here.
 *
 * Store contract: consumes the MailStore API (envelopes, syncState,
 * intents) and never edits store files. Folder purge on full resync goes
 * through the optional SyncStorePort.purgeFolder; stores without it
 * (the P3 InMemoryMailStore) upsert over existing rows, and the T lane
 * SQLite store supplies the transactional purge.
 */

export type ResyncReason = 'uidvalidity-changed' | 'history-expired' | 'delta-expired'

export interface SyncFolderResult {
  decision: SyncDecision
  /** Fresh envelopes to apply (full set on resync, delta otherwise). */
  envelopes: MailEnvelope[]
  /** State to persist after a fully applied page (never partial-mark). */
  state: SyncState
}

export interface SyncDriverPort {
  accountId: string
  kind: AccountKind
  folderId: string
  /** One sync step; the fake decides UIDVALIDITY/history/delta signals. */
  fetchIncremental(stored: SyncState | undefined): Promise<SyncFolderResult>
  /** Transmit with outboxId as the idempotency key. */
  send(req: SendRequest, outboxId: string): Promise<void>
  /** Confirm a possibly-sent id reached the server (Sent reconcile). */
  confirmSent?(outboxId: string): Promise<boolean>
  /** One offline-intent replay attempt against the server. */
  attemptIntent?(intent: MailIntent): Promise<IntentAttempt>
  /** Re-fetch server state for one uid after a stale modseq. */
  refetch?(targetUid: string): Promise<void>
  /** Pause driver timers while offline; resume on reconnect. */
  pause?(): void
  resume?(): void
}

/** MailStore plus the optional folder purge a resync prefers. */
export interface SyncStorePort extends MailStore {
  purgeFolder?(accountId: string, folderId: string): void
}

export interface OfflineStatusPayload {
  online: boolean
  accountId?: string
  detail?: string
}

export interface EngineEmit {
  mailEvent(event: MailEvent): void
  offlineStatus(status: OfflineStatusPayload): void
}

export interface SyncEngineOptions {
  store: SyncStorePort
  clock?: () => string
  emit?: EngineEmit
  makeOutboxId?: () => string
  noTimers?: boolean
}

function folderOf(driver: SyncDriverPort | undefined): string {
  return driver?.folderId ?? ''
}

/** Map engine resync reasons onto the canonical MailEvent resync reasons. */
function toEventReason(reason: ResyncReason): MailEvent & { kind: 'resync' } {
  if (reason === 'uidvalidity-changed') {
    return { kind: 'resync', accountId: '', folderId: '', reason: 'uidvalidity-changed' }
  }
  // historyId and deltaToken are the Gmail/Graph analogues of the IMAP
  // modseq cursor, so an expired cursor surfaces as modseq-expired; the
  // exact cursor kind travels in the companion sync-state detail.
  return { kind: 'resync', accountId: '', folderId: '', reason: 'modseq-expired' }
}

export class SyncEngine {
  private readonly store: SyncStorePort
  private readonly clock: () => string
  private readonly emit?: EngineEmit
  private readonly drivers = new Map<string, SyncDriverPort>()
  private readonly running = new Set<string>()
  private readonly onlineByAccount = new Map<string, boolean>()
  private readonly offlineStatusByAccount = new Map<string, OfflineStatusPayload>()
  private readonly pollFailures = new Map<string, number>()
  private readonly pollTimers = new Map<string, { cancel(): void }>()
  private readonly outbox: OutboxQueue
  private readonly offlineWatchers = new Set<(status: OfflineStatusPayload) => void>()
  private globalOnline = true

  constructor(options: SyncEngineOptions) {
    this.store = options.store
    this.clock = options.clock ?? (() => new Date().toISOString())
    this.emit = options.emit
    const noTimers = options.noTimers ?? false
    this.outbox = new OutboxQueue({
      clock: this.clock,
      makeId: options.makeOutboxId,
      // Deterministic tests pass noTimers and drive drain/replay by hand.
      setTimer: noTimers ? () => ({ cancel() {} }) : undefined,
      onChange: (record) => {
        this.emit?.mailEvent({
          kind: 'sync-state',
          accountId: record.accountId,
          folderId: folderOf(this.drivers.get(record.accountId)),
          state: record.state === 'failed' ? 'error' : record.state === 'sent' ? 'idle' : 'syncing',
          detail: record.lastError,
        })
      },
    })
  }

  registerDriver(port: SyncDriverPort): void {
    this.drivers.set(port.accountId, port)
  }

  private scopeKey(accountId: string, folderId?: string): string {
    return folderId ? `${accountId}/${folderId}` : accountId
  }

  private isOnline(accountId: string): boolean {
    if (!this.globalOnline) return false
    return this.onlineByAccount.get(accountId) ?? true
  }

  /** Current offline status for one account (replayed on subscribe). */
  currentOfflineStatus(accountId?: string): OfflineStatusPayload {
    if (accountId) {
      return (
        this.offlineStatusByAccount.get(accountId) ?? {
          online: this.isOnline(accountId),
          accountId,
        }
      )
    }
    return { online: this.globalOnline }
  }

  /**
   * Main-side offline watcher (used by registerSyncHandlers to fan out to
   * renderers). Watchers fire on every setOnline transition alongside the
   * construction-time emit hook.
   */
  watchOffline(handler: (status: OfflineStatusPayload) => void): () => void {
    this.offlineWatchers.add(handler)
    return () => {
      this.offlineWatchers.delete(handler)
    }
  }

  private announceOffline(status: OfflineStatusPayload): void {
    this.emit?.offlineStatus(status)
    for (const watcher of this.offlineWatchers) watcher(status)
  }

  /**
   * Flip connectivity. Going offline pauses driver timers and emits the
   * status; reconnecting drains outbox sends first, then queued intents.
   */
  async setOnline(online: boolean, accountId?: string, detail?: string): Promise<void> {
    if (accountId) {
      this.onlineByAccount.set(accountId, online)
      const status: OfflineStatusPayload = { online, accountId, detail }
      this.offlineStatusByAccount.set(accountId, status)
      this.announceOffline(status)
      const driver = this.drivers.get(accountId)
      if (online) {
        driver?.resume?.()
        await this.reconnect(accountId)
      } else {
        driver?.pause?.()
      }
      return
    }
    this.globalOnline = online
    const status: OfflineStatusPayload = { online, detail }
    this.announceOffline(status)
    if (online) {
      for (const driver of this.drivers.values()) driver.resume?.()
      for (const driver of this.drivers.values()) await this.reconnect(driver.accountId)
    } else {
      for (const driver of this.drivers.values()) driver.pause?.()
    }
  }

  /**
   * One incremental sync step for an account folder. Debounced per scope:
   * returns started:false when a sync for that scope already runs, or
   * when the account is offline (drivers queue instead of throwing).
   */
  async syncNow(accountId: string, folderId?: string): Promise<{ started: boolean }> {
    const driver = this.drivers.get(accountId)
    if (!driver) throw new Error('mail: unknown account')
    const scope = this.scopeKey(accountId, folderId ?? driver.folderId)
    if (this.running.has(scope)) return { started: false }
    if (!this.isOnline(accountId)) return { started: false }
    const targetFolder = folderId ?? driver.folderId
    this.running.add(scope)
    try {
      this.emit?.mailEvent({
        kind: 'sync-state',
        accountId,
        folderId: targetFolder,
        state: 'syncing',
      })
      const stored = this.store.getSyncState(accountId, targetFolder)
      const result = await driver.fetchIncremental(stored)
      await this.applyResult(driver, targetFolder, stored, result)
      this.pollFailures.set(scope, 0)
      this.emit?.mailEvent({ kind: 'sync-state', accountId, folderId: targetFolder, state: 'idle' })
      return { started: true }
    } catch (err) {
      const failures = (this.pollFailures.get(scope) ?? 0) + 1
      this.pollFailures.set(scope, failures)
      this.emit?.mailEvent({
        kind: 'sync-state',
        accountId,
        folderId: targetFolder,
        state: 'error',
        detail: sanitizeError(err),
      })
      return { started: true }
    } finally {
      this.running.delete(scope)
    }
  }

  private async applyResult(
    driver: SyncDriverPort,
    folderId: string,
    stored: SyncState | undefined,
    result: SyncFolderResult,
  ): Promise<void> {
    void stored
    if (result.decision.kind === 'full-resync') {
      const reason = result.decision.reason as ResyncReason
      // Resync never drops queued outbox rows, drafts, or queued intents:
      // only folder rows are replaced, and only where the store supports
      // a scoped purge (SQLite lane); otherwise fresh rows upsert over
      // stale ones and remarriage happens on uid match.
      this.store.purgeFolder?.(driver.accountId, folderId)
      this.store.upsertEnvelopes(result.envelopes)
      this.store.setSyncState(result.state)
      const event = toEventReason(reason)
      this.emit?.mailEvent({ ...event, accountId: driver.accountId, folderId })
      this.emit?.mailEvent({
        kind: 'sync-state',
        accountId: driver.accountId,
        folderId,
        state: 'idle',
        detail: `resync:${reason}`,
      })
      return
    }
    this.store.upsertEnvelopes(result.envelopes)
    this.store.setSyncState(result.state)
  }

  /** Next poll delay for a scope (pure schedule; failures reset on success). */
  pollDelayFor(accountId: string, folderId?: string, baseMs = POLL_BASE_MS): number {
    const scope = this.scopeKey(accountId, folderId)
    return nextPollDelay({
      baseMs,
      ceilingMs: POLL_CEILING_MS,
      failures: this.pollFailures.get(scope) ?? 0,
    })
  }

  /**
   * Enqueue a send with per-driver caps enforced at enqueue time.
   * Works offline: the entry waits in the outbox until reconnect.
   */
  enqueueSend(accountId: string, req: SendRequest): OutboxRecord {
    const driver = this.drivers.get(accountId)
    if (!driver) throw new Error('mail: unknown account')
    const check = checkOutboxCaps(driver.kind, req)
    if (
      check.verdict === 'too-large' ||
      check.verdict === 'too-many-recipients' ||
      check.verdict === 'attachment-too-large'
    ) {
      throw new OutboxCapError(check.verdict, check.message ?? 'send rejected')
    }
    return this.outbox.enqueue(accountId, req)
  }

  outboxList(accountId: string): OutboxRecord[] {
    return this.outbox.list(accountId)
  }

  outboxStatus(accountId: string): OutboxStatus {
    return this.outbox.status(accountId)
  }

  retryOutbox(accountId: string, outboxId: string): { outboxId: string; state: string } {
    return this.outbox.retry(accountId, outboxId)
  }

  cancelOutbox(accountId: string, outboxId: string): { cancelled: boolean } {
    return this.outbox.cancel(accountId, outboxId)
  }

  /**
   * Drain one account FIFO with a single in-flight send. Caps run again
   * at send time in case they changed while queued. A crash between send
   * and mark resolves via confirmSent before any retransmit.
   */
  async drainOutbox(accountId: string): Promise<void> {
    const driver = this.drivers.get(accountId)
    if (!driver || !this.isOnline(accountId)) return
    // Recover entries stuck in sending by an earlier crash; reconcile
    // below confirms them before any retransmit.
    this.outbox.requeueStuck(accountId)
    for (;;) {
      const next = this.outbox.beginSend(accountId)
      if (!next) break
      // Reconcile first: a previous attempt may have reached the server.
      if (driver.confirmSent) {
        try {
          if (await driver.confirmSent(next.id)) {
            this.outbox.markSent(accountId, next.id)
            continue
          }
        } catch {
          // Confirmation is best-effort; fall through to a fresh send.
        }
      }
      try {
        const check = checkOutboxCaps(driver.kind, next.req)
        if (check.verdict !== 'ok' && check.verdict !== 'warn') {
          throw new OutboxCapError(check.verdict, check.message ?? 'send rejected')
        }
        await driver.send(next.req, next.id)
        this.outbox.markSent(accountId, next.id)
      } catch (err) {
        this.outbox.markAttemptFailed(accountId, next.id, err, () => {
          void this.drainOutbox(accountId)
        })
        // Keep draining later entries only when offline-safe to continue:
        // stop at the first failure so one poisoned entry cannot reorder
        // the queue behind it.
        break
      }
    }
  }

  /** Replay queued offline intents FIFO (server wins on conflict). */
  async replayQueuedIntents(accountId: string): Promise<void> {
    const driver = this.drivers.get(accountId)
    if (!driver || !this.isOnline(accountId)) return
    const port = {
      attemptIntent: (intent: MailIntent): Promise<IntentAttempt> =>
        driver.attemptIntent
          ? driver.attemptIntent(intent)
          : Promise.resolve('acked' as IntentAttempt),
      refetch: (targetUid: string): Promise<void> =>
        driver.refetch?.(targetUid) ?? Promise.resolve(),
    }
    const report = await replayIntents(this.store, port, accountId)
    for (const outcome of report.outcomes) {
      if (outcome.result === 'failed') {
        this.emit?.mailEvent({
          kind: 'conflict',
          accountId,
          folderId: driver.folderId,
          uid: outcome.intentId,
          detail: `intent ${outcome.reason}: server state kept`,
        })
      }
    }
  }

  /** Reconnect order: outbox sends first, then queued intents. */
  async reconnect(accountId: string): Promise<void> {
    await this.drainOutbox(accountId)
    await this.replayQueuedIntents(accountId)
  }

  /** Decide helper re-export for driver authors (pure, tested upstream). */
  decide(signals: SyncSignals) {
    return decideSync(signals)
  }
}
