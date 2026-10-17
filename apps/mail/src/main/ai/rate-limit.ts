/**
 * Token-bucket abuse + cost guard for mail:ai-retrieve (H1 P-9).
 * Per-account buckets: 30 calls/min sustained, burst 10. A call charges
 * one token from EVERY scope member atomically (no partial debit), so a
 * multi-account scope cannot amplify past a single-account budget, and
 * one hot account never starves another. Errors are mail-prefixed and
 * PII-free (429-style `mail: rate-limited, retry in <n>s`). The guard is
 * BYOK-friendly: it bounds call rate, never spend.
 */

export const AI_RETRIEVE_RATE_SUSTAINED_PER_MINUTE = 30

export const AI_RETRIEVE_RATE_BURST = 10

export interface AiRateLimiter {
  /** Charge one token per scope account; throws mail: rate-limited when any member is empty. */
  check(accountIds: string[]): void
  /** Drop all buckets (tests only). */
  reset(): void
}

interface Bucket {
  tokens: number
  updatedMs: number
}

/**
 * Create a token-bucket limiter. The clock is injectable so tests pin
 * burst-then-throttle, refill, and per-account isolation deterministically.
 */
export function createAiRetrieveRateLimiter(now: () => number = Date.now): AiRateLimiter {
  const buckets = new Map<string, Bucket>()

  function refill(bucket: Bucket, atMs: number): void {
    const elapsedMs = Math.max(0, atMs - bucket.updatedMs)
    bucket.tokens = Math.min(
      AI_RETRIEVE_RATE_BURST,
      bucket.tokens + (elapsedMs * AI_RETRIEVE_RATE_SUSTAINED_PER_MINUTE) / 60000,
    )
    bucket.updatedMs = atMs
  }

  return {
    check(accountIds: string[]): void {
      const ids = [...new Set(accountIds)]
      const atMs = now()
      let worstWaitMs = 0
      for (const id of ids) {
        let bucket = buckets.get(id)
        if (!bucket) {
          bucket = { tokens: AI_RETRIEVE_RATE_BURST, updatedMs: atMs }
          buckets.set(id, bucket)
        }
        refill(bucket, atMs)
        if (bucket.tokens < 1) {
          const waitMs = Math.ceil(
            (1 - bucket.tokens) * (60000 / AI_RETRIEVE_RATE_SUSTAINED_PER_MINUTE),
          )
          if (waitMs > worstWaitMs) worstWaitMs = waitMs
        }
      }
      if (worstWaitMs > 0) {
        const retrySec = Math.max(1, Math.ceil(worstWaitMs / 1000))
        throw new Error(`mail: rate-limited, retry in ${retrySec}s`)
      }
      for (const id of ids) {
        const bucket = buckets.get(id)
        if (bucket) bucket.tokens -= 1
      }
    },
    reset(): void {
      buckets.clear()
    },
  }
}

let shared: AiRateLimiter | undefined

/** Process-wide limiter backing mail:ai-retrieve when deps carry none. */
export function sharedAiRetrieveRateLimiter(): AiRateLimiter {
  if (!shared) shared = createAiRetrieveRateLimiter()
  return shared
}

/** Drop all shared buckets (tests only; call in beforeEach). */
export function resetSharedAiRetrieveRateLimiter(): void {
  if (shared) shared.reset()
  else shared = createAiRetrieveRateLimiter()
}
