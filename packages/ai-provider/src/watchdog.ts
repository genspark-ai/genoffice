/**
 * Abort-linked watchdog for LLM requests: a connect timeout until the response
 * headers arrive, then an idle timeout re-armed on every received byte. A
 * silently stalled connection (dropped by a proxy/VPN/firewall without an RST)
 * otherwise leaves the request pending forever and the UI stuck busy.
 */

export const AI_CONNECT_TIMEOUT_MS = 60_000
/**
 * Generous on purpose: on long-context requests the gateway can legitimately go
 * silent for minutes (thinking/buffering before the first token and between
 * chunks). 60s here killed real in-progress generations that were still billed,
 * so only genuinely dead connections should trip this.
 */
export const AI_IDLE_TIMEOUT_MS = 180_000
/** Non-streaming chat waits for the full generation before headers arrive */
export const AI_CHAT_RESPONSE_TIMEOUT_MS = 180_000
/**
 * Floor for watchdog timeouts. Values in (0, floor) are raised to the floor so
 * a tiny misconfigured timeout cannot instant-abort a healthy request.
 */
export const AI_WATCHDOG_MIN_TIMEOUT_MS = 1_000
/**
 * Cap for watchdog timeouts. Infinite or huge values are lowered to the cap so
 * the watchdog still fires instead of going dead. Kept well below the
 * platform setTimeout overflow limit (2^31-1 ms).
 */
export const AI_WATCHDOG_MAX_TIMEOUT_MS = 600_000

/**
 * Normalize a watchdog timeout to a sane finite positive value.
 * Non-numbers, NaN, and values <= 0 fall back to the caller default;
 * Infinity and values above the cap are lowered to the cap;
 * small positives below the floor are raised to the floor.
 */
export function normalizeWatchdogTimeout(ms: number, fallbackMs: number): number {
  if (typeof ms !== 'number' || Number.isNaN(ms) || ms <= 0) return fallbackMs
  if (!Number.isFinite(ms) || ms > AI_WATCHDOG_MAX_TIMEOUT_MS) return AI_WATCHDOG_MAX_TIMEOUT_MS
  if (ms < AI_WATCHDOG_MIN_TIMEOUT_MS) return AI_WATCHDOG_MIN_TIMEOUT_MS
  return ms
}

export class AiTimeoutError extends Error {
  constructor(ms: number) {
    super(`AI request timed out: no data received from the network for ${Math.round(ms / 1000)}s`)
    this.name = 'AiTimeoutError'
  }
}

export interface StreamWatchdog {
  /** pass to fetch: aborts on caller cancel or on timeout */
  signal: AbortSignal
  /** data arrived: switch to (and re-arm) the idle timeout */
  touch(): void
  /** run the request; a timeout abort surfaces as AiTimeoutError; always disposes the timer */
  guard<T>(run: () => Promise<T>): Promise<T>
}

export function createStreamWatchdog(
  parent?: AbortSignal,
  connectMs = AI_CONNECT_TIMEOUT_MS,
  idleMs = AI_IDLE_TIMEOUT_MS,
): StreamWatchdog {
  const normalizedConnectMs = normalizeWatchdogTimeout(connectMs, AI_CONNECT_TIMEOUT_MS)
  const normalizedIdleMs = normalizeWatchdogTimeout(idleMs, AI_IDLE_TIMEOUT_MS)
  const controller = new AbortController()
  let timedOutAfter = 0
  let timer: ReturnType<typeof setTimeout> | undefined
  const arm = (ms: number) => {
    clearTimeout(timer)
    timer = setTimeout(() => {
      timedOutAfter = ms
      controller.abort()
    }, ms)
  }
  const onParentAbort = () => controller.abort()
  if (parent?.aborted) controller.abort()
  else parent?.addEventListener('abort', onParentAbort, { once: true })
  arm(normalizedConnectMs)
  return {
    signal: controller.signal,
    touch: () => arm(normalizedIdleMs),
    async guard<T>(run: () => Promise<T>): Promise<T> {
      try {
        return await run()
      } catch (e) {
        if (timedOutAfter > 0) throw new AiTimeoutError(timedOutAfter)
        throw e
      } finally {
        clearTimeout(timer)
        parent?.removeEventListener('abort', onParentAbort)
      }
    },
  }
}
