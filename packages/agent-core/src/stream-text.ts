import type { AgentTransport } from './types'

/**
 * One tool-less streaming request whose reply body IS the artifact (a page, a
 * document fragment). Text deltas arrive immediately, so the caller can mirror
 * the draft as it grows and no gateway sees a silent connection — tool
 * arguments are buffered server-side until the JSON is complete, and an
 * artifact-sized argument exceeds gateway idle cutoffs.
 */
export interface StreamTextOptions {
  transport: AgentTransport
  system: string
  user: string
  signal?: AbortSignal
  /** the stream is cancelled and reported as partial once the raw reply exceeds this */
  maxChars: number
  /**
   * Maps the raw reply so far to the artifact. `complete` says the format's own
   * terminator arrived (e.g. `</html>`); omit it for formats without one, and
   * the stop reason decides (a natural end is complete, `max_tokens` is not).
   */
  extract(raw: string): { text: string; complete?: boolean }
  /** cumulative extracted text; throttle on the caller's side */
  onProgress?(text: string): void
}

export type StreamTextOutcome =
  | { status: 'complete'; text: string }
  | { status: 'partial'; text: string; reason: 'error' | 'stopped' | 'max_tokens'; error?: string }
  | { status: 'empty'; error: string }

/** Resolves with what arrived, never throws. */
export function streamText(opts: StreamTextOptions): Promise<StreamTextOutcome> {
  return new Promise((resolve) => {
    let raw = ''
    let stopReason: string | undefined
    let settled = false
    const finish = (outcome: StreamTextOutcome) => {
      if (settled) return
      settled = true
      opts.signal?.removeEventListener('abort', onAbort)
      resolve(outcome)
    }
    const partialOrEmpty = (
      reason: 'error' | 'stopped' | 'max_tokens',
      error?: string,
    ): StreamTextOutcome => {
      const { text } = opts.extract(raw)
      if (!text.trim()) return { status: 'empty', error: error ?? reason }
      return error === undefined
        ? { status: 'partial', text, reason }
        : { status: 'partial', text, reason, error }
    }
    const handle = opts.transport.stream(
      { system: opts.system, messages: [{ role: 'user', text: opts.user }], tools: [] },
      {
        onDelta: (delta) => {
          if (settled) return
          raw += delta
          if (raw.length > opts.maxChars) {
            handle.cancel()
            finish(partialOrEmpty('max_tokens', `output exceeded ${opts.maxChars} chars`))
            return
          }
          opts.onProgress?.(opts.extract(raw).text)
        },
        onToolCall: () => undefined,
        onStopReason: (reason) => {
          stopReason = reason
        },
        onDone: () => {
          if (settled) return
          const { text, complete } = opts.extract(raw)
          const finished = complete ?? (stopReason !== 'max_tokens' && text.trim() !== '')
          if (finished) finish({ status: 'complete', text })
          else if (stopReason === 'max_tokens') finish(partialOrEmpty('max_tokens'))
          else finish(partialOrEmpty('error', complete === undefined ? 'empty reply' : undefined))
        },
        onError: (error) => finish(partialOrEmpty('error', error)),
      },
    )
    const onAbort = () => {
      handle.cancel()
      finish(partialOrEmpty('stopped'))
    }
    if (opts.signal?.aborted) onAbort()
    else opts.signal?.addEventListener('abort', onAbort, { once: true })
  })
}
