import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  AI_CONNECT_TIMEOUT_MS,
  AI_IDLE_TIMEOUT_MS,
  AI_WATCHDOG_MAX_TIMEOUT_MS,
  AI_WATCHDOG_MIN_TIMEOUT_MS,
  AiTimeoutError,
  createStreamWatchdog,
  normalizeWatchdogTimeout,
} from '../src/watchdog'

beforeEach(() => {
  vi.useFakeTimers()
})

afterEach(() => {
  vi.useRealTimers()
})

/** promise that rejects when the given signal aborts (models a fetch/body read) */
function abortable(signal: AbortSignal): Promise<never> {
  return new Promise((_, reject) => {
    if (signal.aborted) reject(new Error('aborted'))
    signal.addEventListener('abort', () => reject(new Error('aborted')), { once: true })
  })
}

describe('normalizeWatchdogTimeout', () => {
  it('maps NaN, zero, and negative values to the default', () => {
    expect(normalizeWatchdogTimeout(NaN, AI_CONNECT_TIMEOUT_MS)).toBe(AI_CONNECT_TIMEOUT_MS)
    expect(normalizeWatchdogTimeout(0, AI_CONNECT_TIMEOUT_MS)).toBe(AI_CONNECT_TIMEOUT_MS)
    expect(normalizeWatchdogTimeout(-1, AI_CONNECT_TIMEOUT_MS)).toBe(AI_CONNECT_TIMEOUT_MS)
    expect(normalizeWatchdogTimeout(-Infinity, AI_IDLE_TIMEOUT_MS)).toBe(AI_IDLE_TIMEOUT_MS)
  })

  it('caps Infinity and huge values', () => {
    expect(normalizeWatchdogTimeout(Infinity, AI_CONNECT_TIMEOUT_MS)).toBe(
      AI_WATCHDOG_MAX_TIMEOUT_MS,
    )
    expect(normalizeWatchdogTimeout(Number.MAX_SAFE_INTEGER, AI_CONNECT_TIMEOUT_MS)).toBe(
      AI_WATCHDOG_MAX_TIMEOUT_MS,
    )
  })

  it('floors small positives and leaves normal values unchanged', () => {
    expect(normalizeWatchdogTimeout(1, AI_CONNECT_TIMEOUT_MS)).toBe(AI_WATCHDOG_MIN_TIMEOUT_MS)
    expect(normalizeWatchdogTimeout(999, AI_CONNECT_TIMEOUT_MS)).toBe(AI_WATCHDOG_MIN_TIMEOUT_MS)
    expect(normalizeWatchdogTimeout(1_000, AI_CONNECT_TIMEOUT_MS)).toBe(1_000)
    expect(normalizeWatchdogTimeout(5_000, AI_CONNECT_TIMEOUT_MS)).toBe(5_000)
    expect(normalizeWatchdogTimeout(AI_IDLE_TIMEOUT_MS, AI_IDLE_TIMEOUT_MS)).toBe(
      AI_IDLE_TIMEOUT_MS,
    )
  })
})

describe('watchdog guard with misconfigured timeouts', () => {
  it.each([NaN, 0, -5])(
    'connect=%s degrades to the default instead of aborting instantly',
    async (bad) => {
      const wd = createStreamWatchdog(undefined, bad, 5_000)
      const run = wd.guard(() => abortable(wd.signal))
      const result = expect(run).rejects.toBeInstanceOf(AiTimeoutError)
      // A zero/NaN setTimeout would fire immediately; the guard must stay armed.
      await vi.advanceTimersByTimeAsync(1)
      expect(wd.signal.aborted).toBe(false)
      await vi.advanceTimersByTimeAsync(AI_CONNECT_TIMEOUT_MS - 2)
      expect(wd.signal.aborted).toBe(false)
      await vi.advanceTimersByTimeAsync(1)
      expect(wd.signal.aborted).toBe(true)
      await result
    },
  )

  it('Infinity connect caps instead of going dead', async () => {
    const wd = createStreamWatchdog(undefined, Infinity, 5_000)
    const run = wd.guard(() => abortable(wd.signal))
    const result = expect(run).rejects.toBeInstanceOf(AiTimeoutError)
    await vi.advanceTimersByTimeAsync(AI_WATCHDOG_MAX_TIMEOUT_MS - 1)
    expect(wd.signal.aborted).toBe(false)
    await vi.advanceTimersByTimeAsync(1)
    expect(wd.signal.aborted).toBe(true)
    await result
  })

  it('tiny positive connect floors instead of aborting instantly', async () => {
    const wd = createStreamWatchdog(undefined, 1, 5_000)
    const run = wd.guard(() => abortable(wd.signal))
    const result = expect(run).rejects.toBeInstanceOf(AiTimeoutError)
    await vi.advanceTimersByTimeAsync(1)
    expect(wd.signal.aborted).toBe(false)
    await vi.advanceTimersByTimeAsync(AI_WATCHDOG_MIN_TIMEOUT_MS - 2)
    expect(wd.signal.aborted).toBe(false)
    await vi.advanceTimersByTimeAsync(1)
    expect(wd.signal.aborted).toBe(true)
    await result
  })

  it('bad idle values degrade on touch() instead of aborting instantly or never', async () => {
    for (const bad of [NaN, 0, -10] as const) {
      const wd = createStreamWatchdog(undefined, 60_000, bad)
      const run = wd.guard(() => abortable(wd.signal))
      const result = expect(run).rejects.toBeInstanceOf(AiTimeoutError)
      wd.touch()
      await vi.advanceTimersByTimeAsync(1)
      expect(wd.signal.aborted).toBe(false)
      await vi.advanceTimersByTimeAsync(AI_IDLE_TIMEOUT_MS - 2)
      expect(wd.signal.aborted).toBe(false)
      await vi.advanceTimersByTimeAsync(1)
      expect(wd.signal.aborted).toBe(true)
      await result
    }
  })

  it('Infinity idle caps on touch() instead of going dead', async () => {
    const wd = createStreamWatchdog(undefined, 60_000, Infinity)
    const run = wd.guard(() => abortable(wd.signal))
    const result = expect(run).rejects.toBeInstanceOf(AiTimeoutError)
    wd.touch()
    await vi.advanceTimersByTimeAsync(AI_WATCHDOG_MAX_TIMEOUT_MS - 1)
    expect(wd.signal.aborted).toBe(false)
    await vi.advanceTimersByTimeAsync(1)
    expect(wd.signal.aborted).toBe(true)
    await result
  })

  it('normal values keep their exact schedule', async () => {
    const wd = createStreamWatchdog(undefined, 1_000, 2_000)
    const run = wd.guard(() => abortable(wd.signal))
    const result = expect(run).rejects.toBeInstanceOf(AiTimeoutError)
    await vi.advanceTimersByTimeAsync(999)
    expect(wd.signal.aborted).toBe(false)
    await vi.advanceTimersByTimeAsync(1)
    expect(wd.signal.aborted).toBe(true)
    await result
  })
})
