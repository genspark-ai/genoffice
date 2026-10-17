import { describe, expect, it } from 'vitest'
import {
  OUTBOX_MAX_ATTEMPTS,
  OUTBOX_RETRY_DELAYS_MS,
  POLL_BASE_MS,
  POLL_CEILING_MS,
  nextPollDelay,
  outboxRetryDelay,
} from '../src/main/sync/poll'

describe('poll backoff schedule (spec case: poll-backoff-schedule)', () => {
  it('polls at the 30 s base with no failures', () => {
    expect(nextPollDelay({ baseMs: POLL_BASE_MS, ceilingMs: POLL_CEILING_MS, failures: 0 })).toBe(
      30_000,
    )
  })

  it('doubles per failure with unity jitter', () => {
    const at = (failures: number): number =>
      nextPollDelay({ baseMs: 30_000, ceilingMs: 600_000, failures, random: () => 0.5 })
    expect(at(1)).toBe(60_000)
    expect(at(2)).toBe(120_000)
    expect(at(3)).toBe(240_000)
    expect(at(4)).toBe(480_000)
  })

  it('caps at the 10 minute ceiling', () => {
    const at = (failures: number): number =>
      nextPollDelay({ baseMs: 30_000, ceilingMs: 600_000, failures, random: () => 0.999 })
    expect(at(5)).toBe(600_000)
    expect(at(10)).toBe(600_000)
  })

  it('supports the 30-60 s driver cadence window', () => {
    const delay = nextPollDelay({ baseMs: 45_000, ceilingMs: 600_000, failures: 0 })
    expect(delay).toBe(45_000)
  })
})

describe('outbox retry cadence (spec section 5.3)', () => {
  it('retries at 5 s, 30 s, 2 min, 10 min, max 5 attempts', () => {
    expect(OUTBOX_RETRY_DELAYS_MS).toEqual([5_000, 30_000, 120_000, 600_000])
    expect(OUTBOX_MAX_ATTEMPTS).toBe(5)
    expect(outboxRetryDelay(1)).toBe(5_000)
    expect(outboxRetryDelay(2)).toBe(30_000)
    expect(outboxRetryDelay(3)).toBe(120_000)
    expect(outboxRetryDelay(4)).toBe(600_000)
    expect(outboxRetryDelay(5)).toBeUndefined()
  })
})
