import { afterEach, describe, expect, it, vi } from 'vitest'
import type { AgentToolCall } from '@genoffice/agent-core'
import { retryStreamForProvider } from '../src/stream'
import { okResponse, sseStream } from './test-utils'

afterEach(() => {
  vi.unstubAllGlobals()
})

function collector() {
  const deltas: string[] = []
  const toolCalls: AgentToolCall[] = []
  const stopReasons: string[] = []
  let activityPings = 0
  return {
    deltas,
    toolCalls,
    stopReasons,
    get activityPings() {
      return activityPings
    },
    cb: {
      signal: new AbortController().signal,
      onDelta: (text: string) => deltas.push(text),
      onToolCall: (call: AgentToolCall) => toolCalls.push(call),
      onStopReason: (reason: string) => stopReasons.push(reason),
      onActivity: () => activityPings++,
    },
  }
}

const okBody = () =>
  sseStream([
    'data: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"hello "}}',
    'data: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"world"}}',
  ])

describe('retryStreamForProvider', () => {
  it('retries transient network failures and succeeds on a later attempt', async () => {
    const fetchMock = vi
      .fn()
      .mockRejectedValueOnce(new Error('fetch failed', { cause: { code: 'ECONNRESET' } }))
      .mockResolvedValueOnce(okResponse(okBody()))
    vi.stubGlobal('fetch', fetchMock)
    const { deltas, cb } = collector()
    await retryStreamForProvider(
      'anthropic',
      { apiKey: 'k', model: 'claude-sonnet-5' },
      'sys',
      [],
      [],
      100,
      cb,
    )
    expect(fetchMock).toHaveBeenCalledTimes(2)
    expect(deltas.join('')).toBe('hello world')
  })

  it('retries HTTP 5xx responses but not 4xx', async () => {
    const fiveHundred = new Response('boom', { status: 502 })
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(fiveHundred)
      .mockResolvedValueOnce(okResponse(okBody()))
    vi.stubGlobal('fetch', fetchMock)
    const { deltas, cb } = collector()
    await retryStreamForProvider(
      'anthropic',
      { apiKey: 'k', model: 'claude-sonnet-5' },
      'sys',
      [],
      [],
      100,
      cb,
    )
    expect(fetchMock).toHaveBeenCalledTimes(2)
    expect(deltas.join('')).toBe('hello world')
  })

  it('stops retrying after maxAttempts and rethrows the last error', async () => {
    const fetchMock = vi
      .fn()
      .mockRejectedValue(new Error('fetch failed', { cause: { code: 'ECONNRESET' } }))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await expect(
      retryStreamForProvider(
        'anthropic',
        { apiKey: 'k', model: 'claude-sonnet-5' },
        'sys',
        [],
        [],
        100,
        cb,
        { maxAttempts: 2, backoffMs: [1] },
      ),
    ).rejects.toThrow(/fetch failed/)
    expect(fetchMock).toHaveBeenCalledTimes(2)
  })

  it('does not retry once content has streamed', async () => {
    const encoder = new TextEncoder()
    const partialLine =
      'data: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"partial"}}\n'
    vi.stubGlobal(
      'fetch',
      vi.fn().mockImplementation(() => {
        let pulls = 0
        const stream = new ReadableStream<Uint8Array>({
          pull(controller) {
            pulls++
            if (pulls === 1) controller.enqueue(encoder.encode(partialLine))
            else controller.error(new Error('socket hang up', { cause: { code: 'ECONNRESET' } }))
          },
        })
        return Promise.resolve(new Response(stream, { status: 200 }))
      }),
    )
    const { deltas, cb } = collector()
    await expect(
      retryStreamForProvider(
        'anthropic',
        { apiKey: 'k', model: 'claude-sonnet-5' },
        'sys',
        [],
        [],
        100,
        cb,
        { maxAttempts: 3, backoffMs: [1, 1] },
      ),
    ).rejects.toThrow()
    // the partial text was delivered but the stream never retried
    expect(deltas.join('')).toBe('partial')
    expect(vi.mocked(fetch)).toHaveBeenCalledTimes(1)
  })

  it('re-throws non-transient errors immediately without retry', async () => {
    const fetchMock = vi.fn().mockRejectedValue(new Error('Claude returned no content: x'))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await expect(
      retryStreamForProvider(
        'anthropic',
        { apiKey: 'k', model: 'claude-sonnet-5' },
        'sys',
        [],
        [],
        100,
        cb,
        { maxAttempts: 3, backoffMs: [1, 1] },
      ),
    ).rejects.toThrow(/Claude returned no content/)
    expect(fetchMock).toHaveBeenCalledTimes(1)
  })

  it('backs off between attempts and reports activity so the renderer watchdog stays armed', async () => {
    const fetchMock = vi
      .fn()
      .mockRejectedValueOnce(new Error('fetch failed', { cause: { code: 'ENOTFOUND' } }))
      .mockRejectedValueOnce(new Error('fetch failed', { cause: { code: 'EAI_AGAIN' } }))
      .mockResolvedValueOnce(okResponse(okBody()))
    vi.stubGlobal('fetch', fetchMock)
    const c = collector()
    await retryStreamForProvider(
      'anthropic',
      { apiKey: 'k', model: 'claude-sonnet-5' },
      'sys',
      [],
      [],
      100,
      c.cb,
      { maxAttempts: 3, backoffMs: [5, 5] },
    )
    expect(fetchMock).toHaveBeenCalledTimes(3)
    expect(c.deltas.join('')).toBe('hello world')
    expect(c.activityPings).toBeGreaterThanOrEqual(2)
  })
})
