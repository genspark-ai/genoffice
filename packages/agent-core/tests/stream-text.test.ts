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

describe('streamText', () => {
  it('sends a tool-less request and reports cumulative progress', async () => {
    const progress: string[] = []
    let request: { tools: unknown[]; messages: unknown[] } | undefined
    const outcome = await streamText({
      transport: {
        stream: (req, cb) => {
          request = req
          queueMicrotask(() => {
            cb.onDelta('# Title\n\n')
            cb.onDelta('body')
            cb.onStopReason?.('end_turn')
            cb.onDone()
          })
          return { cancel: () => undefined }
        },
      },
      system: 's',
      user: 'u',
      maxChars: 1000,
      extract: passthrough,
      onProgress: (text) => progress.push(text),
    })
    expect(request?.tools).toEqual([])
    expect(request?.messages).toEqual([{ role: 'user', text: 'u' }])
    expect(outcome).toEqual({ status: 'complete', text: '# Title\n\nbody' })
    expect(progress).toEqual(['# Title', '# Title\n\nbody'])
  })

  it('without a terminator, max_tokens makes the reply partial and end_turn complete', async () => {
    const cut = await streamText({
      transport: fakeTransport((cb) => {
        cb.onDelta('half')
        cb.onStopReason?.('max_tokens')
        cb.onDone()
      }),
      system: 's',
      user: 'u',
      maxChars: 1000,
      extract: passthrough,
    })
    expect(cut).toEqual({ status: 'partial', text: 'half', reason: 'max_tokens' })
    const nothing = await streamText({
      transport: fakeTransport((cb) => {
        cb.onStopReason?.('end_turn')
        cb.onDone()
      }),
      system: 's',
      user: 'u',
      maxChars: 1000,
      extract: passthrough,
    })
    expect(nothing).toEqual({ status: 'empty', error: 'empty reply' })
  })

  it('an explicit terminator decides completeness over the stop reason', async () => {
    const extract = (raw: string) => ({ text: raw, complete: raw.endsWith('END') })
    const done = await streamText({
      transport: fakeTransport((cb) => {
        cb.onDelta('abc END')
        cb.onStopReason?.('max_tokens')
        cb.onDone()
      }),
      system: 's',
      user: 'u',
      maxChars: 1000,
      extract,
    })
    expect(done).toEqual({ status: 'complete', text: 'abc END' })
    const open = await streamText({
      transport: fakeTransport((cb) => {
        cb.onDelta('abc')
        cb.onStopReason?.('end_turn')
        cb.onDone()
      }),
      system: 's',
      user: 'u',
      maxChars: 1000,
      extract,
    })
    expect(open).toEqual({ status: 'partial', text: 'abc', reason: 'error' })
  })

  it('a dropped connection keeps the partial text; an empty one is empty', async () => {
    const partial = await streamText({
      transport: fakeTransport((cb) => {
        cb.onDelta('half')
        cb.onError('connection dropped')
      }),
      system: 's',
      user: 'u',
      maxChars: 1000,
      extract: passthrough,
    })
    expect(partial).toEqual({
      status: 'partial',
      text: 'half',
      reason: 'error',
      error: 'connection dropped',
    })
    const empty = await streamText({
      transport: fakeTransport((cb) => cb.onError('gateway 502')),
      system: 's',
      user: 'u',
      maxChars: 1000,
      extract: passthrough,
    })
    expect(empty).toEqual({ status: 'empty', error: 'gateway 502' })
  })

  it('stop cancels the transport and keeps what arrived; later deltas are ignored', async () => {
    let cancelled = false
    const progress: string[] = []
    const controller = new AbortController()
    let callbacks!: AgentStreamCallbacks
    const promise = streamText({
      transport: {
        stream: (_req, cb) => {
          callbacks = cb
          queueMicrotask(() => cb.onDelta('first'))
          return {
            cancel: () => {
              cancelled = true
            },
          }
        },
      },
      system: 's',
      user: 'u',
      signal: controller.signal,
      maxChars: 1000,
      extract: passthrough,
      onProgress: (text) => progress.push(text),
    })
    await new Promise((r) => setTimeout(r, 0))
    controller.abort()
    callbacks.onDelta(' late')
    callbacks.onDone()
    expect(await promise).toEqual({ status: 'partial', text: 'first', reason: 'stopped' })
    expect(cancelled).toBe(true)
    expect(progress).toEqual(['first'])
  })

  it('the size cap cancels the stream and reports max_tokens', async () => {
    let cancelled = false
    const outcome = await streamText({
      transport: fakeTransport(
        (cb) => cb.onDelta('x'.repeat(20)),
        () => {
          cancelled = true
        },
      ),
      system: 's',
      user: 'u',
      maxChars: 10,
      extract: passthrough,
    })
    expect(outcome).toMatchObject({ status: 'partial', reason: 'max_tokens' })
    expect(cancelled).toBe(true)
  })
})
