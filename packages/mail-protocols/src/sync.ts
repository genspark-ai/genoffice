/**
 * Sync-state machine plus offline intent replay (spec section 4.3).
 * Pure, deterministic, testable without Electron. Conflict rules:
 * server wins for flags/labels/moves; client wins only for queued
 * outbox sends and local draft text. Resync never deletes queued
 * outbox, drafts, or intents (callers preserve those queues).
 */

import type { MailIntent, SyncState } from './types'

export type SyncDecision =
  | { kind: 'initial' }
  | { kind: 'incremental' }
  | { kind: 'full-resync'; reason: 'uidvalidity-changed' | 'history-expired' | 'delta-expired' }

export interface SyncSignals {
  stored: SyncState | undefined
  /** Current server UIDVALIDITY (IMAP). Omitted for Gmail/Graph drivers. */
  remoteUidValidity?: number
  /** Gmail historyId expired (404) since the stored historyId. */
  historyExpired?: boolean
  /** Graph delta token rejected since the stored deltaToken. */
  deltaExpired?: boolean
}

/**
 * Decide the next sync step. UIDVALIDITY change, expired historyId, and
 * expired delta token all force a full folder re-fetch.
 */
export function decideSync(signals: SyncSignals): SyncDecision {
  const { stored } = signals
  if (!stored) return { kind: 'initial' }
  if (signals.remoteUidValidity !== undefined && stored.uidValidity !== undefined) {
    if (signals.remoteUidValidity !== stored.uidValidity) {
      return { kind: 'full-resync', reason: 'uidvalidity-changed' }
    }
  }
  if (signals.historyExpired) return { kind: 'full-resync', reason: 'history-expired' }
  if (signals.deltaExpired) return { kind: 'full-resync', reason: 'delta-expired' }
  return { kind: 'incremental' }
}

/** Intent replay outcome reported by the driver after one attempt. */
export type IntentAttempt = 'acked' | 'stale-modseq' | 'failed'

export interface IntentStep {
  intent: MailIntent
  /** Attempts so far including this step (retry budget: one refetch-retry). */
  attempt: number
}

/**
 * Next replayable intent: oldest queued first. Intents already acked or
 * failed stay out; callers persist state via MailStore.markIntent.
 */
export function nextIntent(
  intents: MailIntent[],
  attempts: Map<string, number>,
): IntentStep | undefined {
  const queued = intents.filter((it) => it.state === 'queued')
  if (queued.length === 0) return undefined
  const intent = queued[0]
  return { intent, attempt: (attempts.get(intent.id) ?? 0) + 1 }
}

export type IntentResolution =
  | { action: 'ack' }
  | { action: 'retry-after-refetch' }
  | { action: 'fail'; reason: 'conflict' | 'transport' }

/**
 * Resolve one attempt: ack wins, stale modseq refetches then retries once
 * and surfaces a conflict on the second stale, transport failure fails.
 * Never silent merge-guesses.
 */
export function resolveIntentAttempt(step: IntentStep, outcome: IntentAttempt): IntentResolution {
  if (outcome === 'acked') return { action: 'ack' }
  if (outcome === 'failed') return { action: 'fail', reason: 'transport' }
  if (step.attempt <= 1) return { action: 'retry-after-refetch' }
  return { action: 'fail', reason: 'conflict' }
}

/**
 * Idempotency key for an intent: same target plus same expected modseq
 * means the same operation (spec section 4.3 item 3).
 */
export function intentKey(accountId: string, targetUid: string, expectedModseq?: string): string {
  return `${accountId}/${targetUid}@${expectedModseq ?? 'none'}`
}
