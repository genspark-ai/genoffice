/**
 * Typed error thrown by AI handlers that fail with a structured code.
 *
 * The IPC wrapper in `apps/web-server/src/index.ts` reads `.message` to
 * surface the failure; we put a human-readable summary there and keep
 * the machine-readable code/channel/reason on the instance for callers
 * that want to branch on them.
 */
export type WebUnsupportedReason =
  | 'renderer-side skill'
  | 'not implemented'
  | 'no upstream provider'

export class WebUnsupportedError extends Error {
  readonly code = 'WEB_UNSUPPORTED' as const
  readonly channel: string
  readonly reason: WebUnsupportedReason

  constructor(channel: string, reason: WebUnsupportedReason = 'renderer-side skill') {
    super(`Channel '${channel}' is not supported on web; ${reason}.`)
    this.name = 'WebUnsupportedError'
    this.channel = channel
    this.reason = reason
  }

  toJSON(): { code: string; channel: string; reason: WebUnsupportedReason } {
    return { code: this.code, channel: this.channel, reason: this.reason }
  }
}

export class InvalidArgumentError extends Error {
  readonly code: 'INVALID_ARGUMENT'
  readonly channel: string
  constructor(channel: string, reason: string) {
    super(`Invalid argument for '${channel}': ${reason}`)
    this.name = 'InvalidArgumentError'
    this.code = 'INVALID_ARGUMENT'
    this.channel = channel
  }
}

export class NotFoundError extends Error {
  readonly code: 'NOT_FOUND'
  readonly channel: string
  constructor(channel: string, reason: string) {
    super(reason)
    this.name = 'NotFoundError'
    this.code = 'NOT_FOUND'
    this.channel = channel
  }
}

export class CorruptError extends Error {
  readonly code: 'CORRUPT'
  readonly channel: string
  constructor(channel: string, reason: string) {
    super(reason)
    this.name = 'CorruptError'
    this.code = 'CORRUPT'
    this.channel = channel
  }
}

/**
 * Best-effort classifier: if the IPC reply is already a structured error
 * code (the renderer emitted one via `throw { code, channel, reason }`),
 * surface the matching class; otherwise wrap the original.
 */
export function classifyWebError(err: unknown, fallbackChannel: string): Error {
  if (err instanceof Error) return err
  if (err && typeof err === 'object') {
    const e = err as { code?: unknown; channel?: unknown; reason?: unknown }
    if (e.code === 'WEB_UNSUPPORTED') {
      return new WebUnsupportedError(
        typeof e.channel === 'string' ? e.channel : fallbackChannel,
        (typeof e.reason === 'string' ? e.reason : 'renderer-side skill') as WebUnsupportedReason,
      )
    }
    if (e.code === 'INVALID_ARGUMENT') {
      return new InvalidArgumentError(
        typeof e.channel === 'string' ? e.channel : fallbackChannel,
        typeof e.reason === 'string' ? e.reason : 'invalid argument',
      )
    }
    if (e.code === 'NOT_FOUND') {
      return new NotFoundError(
        typeof e.channel === 'string' ? e.channel : fallbackChannel,
        typeof e.reason === 'string' ? e.reason : 'not found',
      )
    }
    if (e.code === 'CORRUPT') {
      return new CorruptError(
        typeof e.channel === 'string' ? e.channel : fallbackChannel,
        typeof e.reason === 'string' ? e.reason : 'corrupt data',
      )
    }
  }
  return err instanceof Error ? err : new Error(String(err))
}
