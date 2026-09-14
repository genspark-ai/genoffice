/**
 * Unified AI error model.
 *
 * Every ChatRuntime failure is normalised into an `AIError` with a stable
 * `code` from `AIErrorCode`. The renderer can branch on the code without
 * parsing free-text messages; the user-facing copy is rendered by
 * `AiErrorRecovery` in `@genoffice/ui`.
 *
 * Mirrors the codes surfaced by:
 *   - `apps/web-server/src/ai/errors.ts` (WEB_UNSUPPORTED)
 *   - `packages/agent-core` empty-stream / timeout / provider errors
 *   - XLSX `verifySheetsResponse` failure shape
 */

export type AIErrorCode =
  | 'NOT_CONFIGURED'
  | 'NO_MODEL'
  | 'NETWORK'
  | 'TIMEOUT'
  | 'CANCELLED'
  | 'TOOL_FAILED'
  | 'WEB_UNSUPPORTED'
  | 'PROVIDER'
  | 'INTERNAL'

export class AIError extends Error {
  readonly code: AIErrorCode
  /** Optional channel name (e.g. 'ai:doc-write-continue'). */
  readonly channel?: string | undefined
  /** Optional reason / cause (free-form, machine-readable when possible). */
  readonly cause?: string | undefined
  /** True when the user should be allowed to retry. */
  readonly retryable: boolean

  constructor(
    code: AIErrorCode,
    message: string,
    opts: { channel?: string | undefined; cause?: string | undefined; retryable?: boolean | undefined } = {},
  ) {
    super(message)
    this.name = 'AIError'
    this.code = code
    this.channel = opts.channel
    this.cause = opts.cause
    this.retryable = opts.retryable ?? defaultRetryable(code)
  }
}

function defaultRetryable(code: AIErrorCode): boolean {
  switch (code) {
    case 'NOT_CONFIGURED':
    case 'NO_MODEL':
    case 'WEB_UNSUPPORTED':
      return false
    case 'NETWORK':
    case 'TIMEOUT':
    case 'PROVIDER':
    case 'TOOL_FAILED':
    case 'INTERNAL':
      return true
    case 'CANCELLED':
      return false
  }
}

/**
 * Best-effort classifier: walks common error shapes (Error, AIError,
 * IPC `{code,channel,reason}` envelopes, plain strings) and returns an
 * `AIError`. The fallback code is `INTERNAL`.
 */
export function classifyError(err: unknown, fallbackChannel?: string): AIError {
  if (err instanceof AIError) return err

  if (typeof err === 'string') {
    const trimmed = err.trim()
    if (trimmed.startsWith('{') && trimmed.endsWith('}')) {
      try {
        return classifyError(JSON.parse(trimmed), fallbackChannel)
      } catch {
        /* not JSON — fall through to the plain string branch below */
      }
    }
    // Mirror the Error branch's regex order so plain-string inputs
    // (AgentLoop's onError delivers them as raw strings) get the same
    // WEB_UNSUPPORTED / NETWORK / PROVIDER classification instead of
    // collapsing to INTERNAL.
    if (/HTTP 403.*web page|HTTP 401.*sign in|sign.in.to.use|sign.in.required/i.test(trimmed)) {
      return new AIError('WEB_UNSUPPORTED', err, { channel: fallbackChannel })
    }
    if (/(timeout|timed out)/i.test(trimmed)) {
      return new AIError('TIMEOUT', err, { channel: fallbackChannel })
    }
    if (/(network|ENOTFOUND|ECONNRESET|ECONNREFUSED|ETIMEDOUT|fetch failed)/i.test(trimmed)) {
      return new AIError('NETWORK', err, { channel: fallbackChannel })
    }
    if (/No API key|No model selected|not configured/i.test(trimmed)) {
      return new AIError('NOT_CONFIGURED', err, { channel: fallbackChannel })
    }
    if (/rate.limit|quota|usage.limit|overloaded|too.many.requests|HTTP 429|用量上限/i.test(trimmed)) {
      return new AIError('PROVIDER', err, { channel: fallbackChannel })
    }
    if (/cancel/i.test(trimmed)) {
      return new AIError('CANCELLED', err, { channel: fallbackChannel })
    }
    return new AIError('INTERNAL', err, { channel: fallbackChannel })
  }

  if (err instanceof Error) {
    const message = err.message
    if (/^Channel '.*' is not supported on web/i.test(message)) {
      return new AIError('WEB_UNSUPPORTED', message, { channel: fallbackChannel })
    }
    if (/HTTP 403.*web page|HTTP 401.*sign in|sign.in.to.use|sign.in.required/i.test(message)) {
      // The provider's auth gate returned an HTML page instead of an API
      // response (Genspark proxy, gsk-logged-out flows, OAuth-only vendors).
      // Surface this as WEB_UNSUPPORTED so the user knows the chosen
      // provider requires a desktop / signed-in session, not a retry.
      return new AIError('WEB_UNSUPPORTED', message, { channel: fallbackChannel })
    }
    if (/(timeout|timed out)/i.test(message)) {
      return new AIError('TIMEOUT', message, { channel: fallbackChannel })
    }
    if (/(network|ENOTFOUND|ECONNRESET|ECONNREFUSED|ETIMEDOUT|fetch failed)/i.test(message)) {
      return new AIError('NETWORK', message, { channel: fallbackChannel })
    }
    if (/No API key|No model selected|not configured/i.test(message)) {
      return new AIError('NOT_CONFIGURED', message, { channel: fallbackChannel })
    }
    if (/rate.limit|quota|usage.limit|overloaded|too.many.requests|HTTP 429|用量上限/i.test(message)) {
      return new AIError('PROVIDER', message, { channel: fallbackChannel })
    }
    if (/cancel/i.test(message)) {
      return new AIError('CANCELLED', message, { channel: fallbackChannel })
    }
    return new AIError('INTERNAL', message, { channel: fallbackChannel, cause: err.stack })
  }

  if (err && typeof err === 'object') {
    const e = err as { code?: unknown; channel?: unknown; reason?: unknown; message?: unknown }
    if (e.code === 'WEB_UNSUPPORTED') {
      return new AIError(
        'WEB_UNSUPPORTED',
        typeof e.message === 'string' ? e.message : `Channel not supported on web: ${String(e.channel ?? fallbackChannel ?? '')}`,
        { channel: typeof e.channel === 'string' ? e.channel : fallbackChannel, cause: typeof e.reason === 'string' ? e.reason : undefined },
      )
    }
    if (typeof e.code === 'string' && isAIErrorCode(e.code)) {
      return new AIError(
        e.code,
        typeof e.message === 'string' ? e.message : e.code,
        {
          channel: typeof e.channel === 'string' ? e.channel : fallbackChannel,
          cause: typeof e.reason === 'string' ? e.reason : undefined,
        },
      )
    }
  }

  return new AIError('INTERNAL', String(err), { channel: fallbackChannel })
}

function isAIErrorCode(s: string): s is AIErrorCode {
  return [
    'NOT_CONFIGURED',
    'NO_MODEL',
    'NETWORK',
    'TIMEOUT',
    'CANCELLED',
    'TOOL_FAILED',
    'WEB_UNSUPPORTED',
    'PROVIDER',
    'INTERNAL',
  ].includes(s)
}

export function isAIError(err: unknown): err is AIError {
  return err instanceof AIError
}
