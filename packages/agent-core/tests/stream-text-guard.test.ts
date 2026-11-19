import { describe, expect, it } from 'vitest'
import type { AgentStreamCallbacks, AgentTransport } from '../src/types'
import { streamText } from '../src/stream-text'

function fakeTransport(
  script: (cb: AgentStreamCallbacks) => void,
  onCancel?: () => void,
): AgentTransport {
  return {
    stream: (_req, cb) => {
      queueMicrotask(() => script(cb))
      return { cancel: () => onCancel?.() }
    },
  }
}

const passthrough = (raw: string) => ({ text: raw.trim() })

// Fail fast instead of hanging until the vitest timeout.
function withTimeout<T>(promise: Promise<T>, ms = 1000): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined
  const timeout = new Promise<never>((_, reject) => {
    timer = setTimeout(
      () => reject(new Error(`timed out after ${ms}ms (promise never settled)`)),
      ms,
    )
  })
  return Promise.race([promise, timeout]).finally(() => clearTimeout(timer))
}

describe('streamText guard', () => {
  it('NaN maxChars falls back to the default cap instead of never firing', async () => {
    // Small reply stays complete under the fallback cap.
    const small = await withTimeout(
      streamText({
        transport: fakeTransport((cb) => {
          cb.onDelta('hello')
          cb.onStopReason?.('end_turn')
          cb.onDone()
        }),
        system: 's',
        user: 'u',
        maxChars: NaN,
        extract: passthrough,
      }),
    )
    expect(small).toEqual({ status: 'complete', text: 'hello' })

    // Reply beyond the fallback cap is cut off as max_tokens.
    let cancelled = false
    const big = await withTimeout(
      streamText({
        transport: fakeTransport(
          (cb) => cb.onDelta('x'.repeat(200001)),
          () => {
            cancelled = true
          },
        ),
        system: 's',
        user: 'u',
        maxChars: NaN,
        extract: passthrough,
      }),
    )
    expect(big).toMatchObject({ status: 'partial', reason: 'max_tokens' })
    expect(cancelled).toBe(true)
  })

  it('non-positive maxChars does not trip immediately', async () => {
    for (const maxChars of [0, -10]) {
      const outcome = await withTimeout(
        streamText({
          transport: fakeTransport((cb) => {
            cb.onDelta('hello')
            cb.onStopReason?.('end_turn')
            cb.onDone()
          }),
          system: 's',
          user: 'u',
          maxChars,
          extract: passthrough,
        }),
      )
      expect(outcome).toEqual({ status: 'complete', text: 'hello' })
    }
  })

  it('throwing extract resolves with an error result instead of hanging', async () => {
    const throwing = () => {
      throw new Error('extract boom')
    }

    // Partial text is kept alongside the extract error.
    const partial = await withTimeout(
      streamText({
        transport: fakeTransport((cb) => {
          cb.onDelta('half')
          cb.onStopReason?.('end_turn')
          cb.onDone()
        }),
        system: 's',
        user: 'u',
        maxChars: 1000,
        extract: throwing,
      }),
    )
    expect(partial).toEqual({
      status: 'partial',
      text: 'half',
      reason: 'error',
      error: 'extract boom',
    })

    // Empty reply surfaces as empty with the extract error.
    const empty = await withTimeout(
      streamText({
        transport: fakeTransport((cb) => {
          cb.onStopReason?.('end_turn')
          cb.onDone()
        }),
        system: 's',
        user: 'u',
        maxChars: 1000,
        extract: throwing,
      }),
    )
    expect(empty).toEqual({ status: 'empty', error: 'extract boom' })
  })

  it('throwing extract on progress still settles', async () => {
    // Fails only for the first progress update; the final reply extracts fine.
    const flaky = (raw: string) => {
      if (raw === 'first') throw new Error('progress boom')
      return { text: raw.trim() }
    }
    const progress: string[] = []
    const outcome = await withTimeout(
      streamText({
        transport: fakeTransport((cb) => {
          cb.onDelta('first')
          cb.onDelta(' second')
          cb.onStopReason?.('end_turn')
          cb.onDone()
        }),
        system: 's',
        user: 'u',
        maxChars: 1000,
        extract: flaky,
        onProgress: (text) => progress.push(text),
      }),
    )
    expect(outcome).toEqual({ status: 'complete', text: 'first second' })
    expect(progress).toEqual(['first', 'first second'])
  })

  it('normal path is unchanged', async () => {
    const outcome = await withTimeout(
      streamText({
        transport: fakeTransport((cb) => {
          cb.onDelta('abc')
          cb.onStopReason?.('end_turn')
          cb.onDone()
        }),
        system: 's',
        user: 'u',
        maxChars: 1000,
        extract: passthrough,
      }),
    )
    expect(outcome).toEqual({ status: 'complete', text: 'abc' })
  })
})
