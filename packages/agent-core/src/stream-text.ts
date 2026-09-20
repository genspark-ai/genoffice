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
    // Fallback cap when the caller passes NaN, Infinity, zero, or a negative limit.
    const FALLBACK_MAX_CHARS = 200000
    const maxChars =
      Number.isFinite(opts.maxChars) && opts.maxChars > 0 ? opts.maxChars : FALLBACK_MAX_CHARS
    // Extract never throws; on failure the raw reply is kept so the promise settles.
    const safeExtract = (input: string): { text: string; complete?: boolean; error?: string } => {
      try {
        return opts.extract(input)
      } catch (err) {
        return { text: input, error: err instanceof Error ? err.message : String(err) }
      }
    }
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
      const extracted = safeExtract(raw)
      const resolvedError = error ?? extracted.error
      if (!extracted.text.trim()) return { status: 'empty', error: resolvedError ?? reason }
      return resolvedError === undefined
        ? { status: 'partial', text: extracted.text, reason }
        : { status: 'partial', text: extracted.text, reason, error: resolvedError }
    }
    const handle = opts.transport.stream(
      { system: opts.system, messages: [{ role: 'user', text: opts.user }], tools: [] },
      {
        onDelta: (delta) => {
          if (settled) return
          raw += delta
          if (raw.length > maxChars) {
            handle.cancel()
            finish(partialOrEmpty('max_tokens', `output exceeded ${maxChars} chars`))
            return
          }
          try {
            opts.onProgress?.(safeExtract(raw).text)
          } catch {
            // Ignore progress listener failures so the stream can still settle.
          }
        },
        onToolCall: () => undefined,
        onStopReason: (reason) => {
          stopReason = reason
        },
        onDone: () => {
          if (settled) return
          const { text, complete, error: extractError } = safeExtract(raw)
          const finished =
            extractError === undefined &&
            (complete ?? (stopReason !== 'max_tokens' && text.trim() !== ''))
          if (finished) finish({ status: 'complete', text })
          else if (stopReason === 'max_tokens') finish(partialOrEmpty('max_tokens', extractError))
          else
            finish(
              partialOrEmpty(
                'error',
                extractError ?? (complete === undefined ? 'empty reply' : undefined),
              ),
            )
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
