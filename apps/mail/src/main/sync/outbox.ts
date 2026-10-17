import type { SendRequest } from '@genoffice/mail-protocols'
import { OUTBOX_MAX_ATTEMPTS, outboxRetryDelay } from './poll'
import { sanitizeError } from './redact'

/**
 * Outbox replay queue (P4 spec section 5). FIFO per account, one in-flight
 * send per account, automatic retry at 5 s / 30 s / 2 min / 10 min up to
 * 5 attempts, then failed. Manual retry resets the budget and re-queues
 * at head. Cancelled entries are never sent, even mid-retry.
 *
 * Idempotency: every entry carries a stable id that doubles as the
 * idempotency key. Drivers send with `Message-ID:
 * <outbox-<id>@genoffice.local>` (SMTP) or a check-then-send inside one
 * marking step (Gmail/Graph), so a crash between send and mark resolves
 * to at-most-once via the Sent reconcile on next sync.
 */

export type OutboxState = 'queued' | 'sending' | 'failed' | 'sent' | 'cancelled'

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

export interface OutboxStatus {
  queued: number
  failed: number
}

export interface OutboxQueueOptions {
  /** ISO clock; injectable for deterministic tests. */
  clock?: () => string
  /** Id factory; injectable for deterministic tests. */
  makeId?: () => string
  /** Timer factory; injectable so tests avoid real timers. */
  setTimer?: (delayMs: number, fn: () => void) => { cancel(): void }
  /** Called on every state transition (engine forwards to MailEvents). */
  onChange?: (record: OutboxRecord) => void
}

interface QueueEntry extends OutboxRecord {
  /** Manual retry parks the entry ahead of FIFO order. */
  head: boolean
  timer?: { cancel(): void }
}

function defaultTimer(delayMs: number, fn: () => void): { cancel(): void } {
  const handle = setTimeout(fn, delayMs)
  if (typeof (handle as { unref?: unknown }).unref === 'function') {
    ;(handle as { unref(): void }).unref()
  }
  return { cancel: () => clearTimeout(handle) }
}

/** Thrown when a send request breaks a per-driver cap at enqueue time. */
export class OutboxCapError extends Error {
  readonly verdict: string
  constructor(verdict: string, message: string) {
    super(`outbox rejected (${verdict}): ${message}`)
    this.name = 'OutboxCapError'
    this.verdict = verdict
  }
}

export class OutboxQueue {
  private readonly entries = new Map<string, QueueEntry>()
  private readonly inFlight = new Set<string>()
  private readonly clock: () => string
  private readonly makeId: () => string
  private readonly setTimer: (delayMs: number, fn: () => void) => { cancel(): void }
  private readonly onChange?: (record: OutboxRecord) => void
  private nextId = 1

  constructor(options?: OutboxQueueOptions) {
    this.clock = options?.clock ?? (() => new Date().toISOString())
    this.makeId = options?.makeId ?? (() => `outbox-${this.nextId++}`)
    this.setTimer = options?.setTimer ?? defaultTimer
    this.onChange = options?.onChange
  }

  private snapshot(entry: QueueEntry): OutboxRecord {
    const { head: _head, timer: _timer, ...record } = entry
    return { ...record }
  }

  private changed(entry: QueueEntry): void {
    this.onChange?.(this.snapshot(entry))
  }

  /** Enqueue a send; caps are checked by the caller before this lands. */
  enqueue(accountId: string, req: SendRequest): OutboxRecord {
    const now = this.clock()
    const entry: QueueEntry = {
      id: this.makeId(),
      accountId,
      req,
      state: 'queued',
      attempts: 0,
      queuedAt: now,
      head: false,
    }
    this.entries.set(entry.id, entry)
    this.changed(entry)
    return this.snapshot(entry)
  }

  get(accountId: string, outboxId: string): OutboxRecord | undefined {
    const entry = this.entries.get(outboxId)
    if (!entry || entry.accountId !== accountId) return undefined
    return this.snapshot(entry)
  }

  /** Newest-first, capped at 200 rows (O lane contract). */
  list(accountId: string): OutboxRecord[] {
    return [...this.entries.values()]
      .filter((entry) => entry.accountId === accountId)
      .sort((a, b) => (a.queuedAt < b.queuedAt ? 1 : a.queuedAt > b.queuedAt ? -1 : 0))
      .slice(0, 200)
      .map((entry) => this.snapshot(entry))
  }

  status(accountId: string): OutboxStatus {
    let queued = 0
    let failed = 0
    for (const entry of this.entries.values()) {
      if (entry.accountId !== accountId) continue
      if (entry.state === 'queued' || entry.state === 'sending') queued += 1
      if (entry.state === 'failed') failed += 1
    }
    return { queued, failed }
  }

  /** FIFO order for one account: head-priority first, then queuedAt, then id. */
  private fifo(accountId: string): QueueEntry[] {
    return [...this.entries.values()]
      .filter((entry) => entry.accountId === accountId && entry.state === 'queued')
      .sort((a, b) => {
        if (a.head !== b.head) return a.head ? -1 : 1
        if (a.queuedAt !== b.queuedAt) return a.queuedAt < b.queuedAt ? -1 : 1
        return a.id < b.id ? -1 : 1
      })
  }

  /** Next queued entry, or undefined when the account has nothing to send. */
  nextQueued(accountId: string): OutboxRecord | undefined {
    const next = this.fifo(accountId)[0]
    return next ? this.snapshot(next) : undefined
  }

  isBusy(accountId: string): boolean {
    return this.inFlight.has(accountId)
  }

  /**
   * One send attempt for the head entry. The caller transmits with
   * entry.id as the idempotency key, then settles with markSent or
   * markAttemptFailed. Returns undefined when busy, offline, or empty.
   */
  beginSend(accountId: string): OutboxRecord | undefined {
    if (this.inFlight.has(accountId)) return undefined
    const next = this.fifo(accountId)[0]
    if (!next) return undefined
    if (next.state !== 'queued') return undefined
    this.inFlight.add(accountId)
    next.state = 'sending'
    next.attempts += 1
    next.head = false
    this.changed(next)
    return this.snapshot(next)
  }

  markSent(accountId: string, outboxId: string): OutboxRecord | undefined {
    const entry = this.entries.get(outboxId)
    if (!entry || entry.accountId !== accountId) return undefined
    entry.timer?.cancel()
    entry.timer = undefined
    entry.state = 'sent'
    entry.sentAt = this.clock()
    entry.nextRetryAt = undefined
    entry.lastError = undefined
    this.inFlight.delete(accountId)
    this.changed(entry)
    return this.snapshot(entry)
  }

  /**
   * Record a failed attempt: schedule the next retry while budget
   * remains, else park the entry as failed for the UI + manual retry.
   */
  markAttemptFailed(
    accountId: string,
    outboxId: string,
    err: unknown,
    onRetryDue?: () => void,
  ): OutboxRecord | undefined {
    const entry = this.entries.get(outboxId)
    if (!entry || entry.accountId !== accountId) return undefined
    this.inFlight.delete(accountId)
    entry.lastError = sanitizeError(err)
    if (entry.attempts >= OUTBOX_MAX_ATTEMPTS) {
      entry.state = 'failed'
      entry.nextRetryAt = undefined
      this.changed(entry)
      return this.snapshot(entry)
    }
    const delay = outboxRetryDelay(entry.attempts) ?? 600_000
    entry.state = 'queued'
    entry.nextRetryAt = new Date(Date.parse(this.clock()) + delay).toISOString()
    entry.timer?.cancel()
    if (onRetryDue) entry.timer = this.setTimer(delay, onRetryDue)
    this.changed(entry)
    return this.snapshot(entry)
  }

  /**
   * Reconcile a send that may have reached the server before a crash:
   * when the server confirms the id, mark sent without retransmitting.
   * Returns true when the entry resolved to sent.
   */
  reconcileConfirmed(accountId: string, outboxId: string): boolean {
    const entry = this.entries.get(outboxId)
    if (!entry || entry.accountId !== accountId) return false
    if (entry.state === 'sent') return true
    if (entry.state === 'cancelled') return false
    this.markSent(accountId, outboxId)
    return true
  }

  /** Manual retry: reset the budget and re-queue at head. */
  retry(accountId: string, outboxId: string): { outboxId: string; state: OutboxState } {
    const entry = this.entries.get(outboxId)
    if (!entry || entry.accountId !== accountId) throw new Error('mail: unknown outbox entry')
    if (entry.state === 'sent' || entry.state === 'sending') {
      throw new Error('mail: entry is already sending or sent')
    }
    entry.timer?.cancel()
    entry.timer = undefined
    entry.state = 'queued'
    entry.attempts = 0
    entry.nextRetryAt = undefined
    entry.lastError = undefined
    entry.head = true
    this.changed(entry)
    return { outboxId: entry.id, state: 'queued' }
  }

  /** Cancel a queued or failed entry; sending/sent entries refuse. */
  cancel(accountId: string, outboxId: string): { cancelled: boolean } {
    const entry = this.entries.get(outboxId)
    if (!entry || entry.accountId !== accountId) return { cancelled: false }
    if (entry.state === 'sent' || entry.state === 'sending') return { cancelled: false }
    if (entry.state === 'cancelled') return { cancelled: true }
    entry.timer?.cancel()
    entry.timer = undefined
    entry.state = 'cancelled'
    entry.nextRetryAt = undefined
    entry.head = false
    this.changed(entry)
    return { cancelled: true }
  }

  /** Drop sent entries older than 7 days (kept for Sent confirmation). */
  pruneSent(olderThanDays = 7): number {
    const cutoff = Date.parse(this.clock()) - olderThanDays * 24 * 3600_000
    let pruned = 0
    for (const [id, entry] of this.entries) {
      if (entry.state !== 'sent') continue
      if (Date.parse(entry.sentAt ?? entry.queuedAt) < cutoff) {
        this.entries.delete(id)
        pruned += 1
      }
    }
    return pruned
  }

  /** Test seam: release one in-flight account without settling the entry. */
  release(accountId: string): void {
    this.inFlight.delete(accountId)
  }

  /**
   * Requeue entries stuck in sending (process restart between send and
   * mark). Attempts are kept; the drain reconciles via confirmSent before
   * any retransmit, so recovery stays at-most-once.
   */
  requeueStuck(accountId: string): number {
    let moved = 0
    for (const entry of this.entries.values()) {
      if (entry.accountId !== accountId || entry.state !== 'sending') continue
      entry.state = 'queued'
      entry.nextRetryAt = undefined
      moved += 1
    }
    this.inFlight.delete(accountId)
    return moved
  }
}
