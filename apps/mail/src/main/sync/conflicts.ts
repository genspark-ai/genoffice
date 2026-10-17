import {
  nextIntent,
  resolveIntentAttempt,
  type IntentAttempt,
  type MailIntent,
} from '@genoffice/mail-protocols'

/**
 * Server-wins conflict wiring (P4 spec section 4). The pure decision
 * helpers (nextIntent, resolveIntentAttempt, intentKey) live in
 * mail-protocols sync.ts and are reused here, never rewritten.
 *
 * Server wins for flags, labels, and moves; the client wins only for
 * queued outbox sends and local draft text. Offline moves/flags replay
 * FIFO on reconnect; a stale modseq refetches once and retries once,
 * then surfaces a conflict event instead of guessing a merge.
 */

export interface IntentDriver {
  /** One replay attempt against the server. */
  attemptIntent(intent: MailIntent): Promise<IntentAttempt>
  /** Re-fetch server state for one uid after a stale modseq. */
  refetch?(targetUid: string): Promise<void>
}

export interface IntentStore {
  pendingIntents(accountId: string): MailIntent[]
  markIntent(id: string, state: 'acked' | 'failed'): void
}

export type IntentOutcome =
  | { intentId: string; result: 'acked' }
  | { intentId: string; result: 'failed'; reason: 'conflict' | 'transport' }

export interface ReplayReport {
  outcomes: IntentOutcome[]
}

/**
 * Replay every queued intent FIFO. Stale-modseq attempts refetch once
 * and retry once; a second stale or a transport failure marks the intent
 * failed so the caller can emit the conflict event. Never merges.
 */
export async function replayIntents(
  store: IntentStore,
  driver: IntentDriver,
  accountId: string,
): Promise<ReplayReport> {
  const attempts = new Map<string, number>()
  const outcomes: IntentOutcome[] = []
  const guard = store.pendingIntents(accountId).length * 3 + 10
  for (let step = 0; step < guard; step += 1) {
    const queued = store.pendingIntents(accountId)
    const next = nextIntent(queued, attempts)
    if (!next) break
    attempts.set(next.intent.id, next.attempt)
    const outcome = await driver.attemptIntent(next.intent)
    const resolution = resolveIntentAttempt(next, outcome)
    if (resolution.action === 'ack') {
      store.markIntent(next.intent.id, 'acked')
      outcomes.push({ intentId: next.intent.id, result: 'acked' })
      continue
    }
    if (resolution.action === 'retry-after-refetch') {
      await driver.refetch?.(next.intent.targetUid)
      continue
    }
    store.markIntent(next.intent.id, 'failed')
    outcomes.push({ intentId: next.intent.id, result: 'failed', reason: resolution.reason })
  }
  return { outcomes }
}
