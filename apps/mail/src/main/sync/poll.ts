/**
 * Poll/backoff schedule for sync drivers (P4 spec sections 3.1-3.3).
 * IMAP polls at 30 s when IDLE is not advertised; Gmail and Graph poll
 * history/delta links every 30-60 s. Failures back off exponentially with
 * jitter to a 10 minute ceiling; success resets to the base cadence.
 * Pure and deterministic so tests pin the schedule without timers.
 */

export const POLL_BASE_MS = 30_000
export const POLL_CEILING_MS = 10 * 60_000

export interface PollSchedule {
  /** Base cadence in ms (default 30 s). */
  baseMs: number
  /** Backoff ceiling in ms (default 10 min). */
  ceilingMs: number
  /** Consecutive failures since the last success. */
  failures: number
  /** Random source in [0, 1); injectable for deterministic tests. */
  random?: () => number
}

/**
 * Delay before the next poll: base on success, doubling per failure with
 * up to 20 percent jitter, capped at the ceiling.
 */
export function nextPollDelay(schedule: PollSchedule): number {
  const baseMs = schedule.baseMs > 0 ? schedule.baseMs : POLL_BASE_MS
  const ceilingMs = schedule.ceilingMs > 0 ? schedule.ceilingMs : POLL_CEILING_MS
  const failures = Math.max(0, Math.floor(schedule.failures))
  if (failures <= 0) return Math.min(baseMs, ceilingMs)
  const grown = baseMs * 2 ** failures
  const rand = schedule.random ? schedule.random() : Math.random()
  const jittered = grown * (0.9 + 0.2 * Math.min(Math.max(rand, 0), 1))
  return Math.min(Math.floor(jittered), ceilingMs)
}

/** Retry cadence for outbox sends: 5 s, 30 s, 2 min, 10 min, then stop. */
export const OUTBOX_RETRY_DELAYS_MS = [5_000, 30_000, 120_000, 600_000]

/** Maximum automatic send attempts before an entry goes failed. */
export const OUTBOX_MAX_ATTEMPTS = 5

/** Delay before the nth retry (1-based); undefined when no retry remains. */
export function outboxRetryDelay(attempts: number): number | undefined {
  if (attempts < 1 || attempts > OUTBOX_RETRY_DELAYS_MS.length) return undefined
  return OUTBOX_RETRY_DELAYS_MS[attempts - 1]
}
