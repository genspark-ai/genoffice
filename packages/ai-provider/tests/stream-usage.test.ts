import { afterEach, describe, expect, it, vi } from 'vitest'
import type { AgentToolCall } from '@genoffice/agent-core'
import { streamForProvider } from '../src/stream'
import { resetUsageRejectedGateways } from '../src/protocols/openai-compatible'
import type { AiTokenUsage } from '../src/types'
import { errorResponse, jsonResponse, okResponse, sseStream } from './test-utils'

afterEach(() => {
  vi.unstubAllGlobals()
  resetUsageRejectedGateways()
})

function collector() {
  const deltas: string[] = []
  const toolCalls: AgentToolCall[] = []
  const stopReasons: string[] = []
  const usages: AiTokenUsage[] = []
  return {
    deltas,
    toolCalls,
    stopReasons,
    usages,
    cb: {
      signal: new AbortController().signal,
      onDelta: (text: string) => deltas.push(text),
      onToolCall: (call: AgentToolCall) => toolCalls.push(call),
      onStopReason: (reason: string) => stopReasons.push(reason),
      onUsage: (usage: AiTokenUsage) => usages.push(usage),
    },
  }
}

async function run(provider: 'openai' | 'anthropic' | 'gemini', c: ReturnType<typeof collector>) {
  await streamForProvider(provider, { apiKey: 'k', model: 'm' }, 'sys', [], [], 1024, c.cb)
}

describe('openai-compatible usage', () => {
  it('requests include_usage and reports the usage chunk (empty choices)', async () => {
    const requests: Array<Record<string, unknown>> = []
    vi.stubGlobal(
      'fetch',
      vi.fn((_url: unknown, init?: { body?: string }) => {
        requests.push(init?.body ? (JSON.parse(init.body) as Record<string, unknown>) : {})
        return okResponse(
          sseStream([
            'data: {"choices":[{"delta":{"content":"hi"}}]}',
            'data: {"choices":[],"usage":{"prompt_tokens":123,"completion_tokens":45}}',
            'data: [DONE]',
            '',
          ]),
        )
      }),
    )
    const c = collector()
    await run('openai', c)
    expect(c.deltas.join('')).toBe('hi')
    expect(c.usages).toEqual([{ inputTokens: 123, outputTokens: 45 }])
    expect(requests[0]!.stream_options).toEqual({ include_usage: true })
  })

  it('accepts usage on the final chunk with empty delta (DeepSeek-style)', async () => {
    vi.stubGlobal(
      'fetch',
      vi
        .fn()
        .mockResolvedValue(
          okResponse(
            sseStream([
              'data: {"choices":[{"delta":{"content":"hi"}}]}',
              'data: {"choices":[{"delta":{},"finish_reason":"stop"}],"usage":{"prompt_tokens":"7","completion_tokens":"3"}}',
              'data: [DONE]',
              '',
            ]),
          ),
        ),
    )
    const c = collector()
    await run('openai', c)
    expect(c.usages).toEqual([{ inputTokens: 7, outputTokens: 3 }])
  })

  it('retries once without stream_options when the gateway rejects the field', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(errorResponse(400, 'Unknown parameter: stream_options'))
      .mockImplementation(async () =>
        okResponse(
          sseStream(['data: {"choices":[{"delta":{"content":"ok"}}]}', 'data: [DONE]', '']),
        ),
      )
    vi.stubGlobal('fetch', fetchMock)
    const c = collector()
    await run('openai', c)
    expect(fetchMock).toHaveBeenCalledTimes(2)
    const secondBody = JSON.parse((fetchMock.mock.calls[1]![1] as { body: string }).body) as Record<
      string,
      unknown
    >
    expect(secondBody.stream_options).toBeUndefined()
    expect(c.deltas.join('')).toBe('ok')

    // the rejection is remembered: the next turn skips the field and the retry
    await run('openai', collector())
    expect(fetchMock).toHaveBeenCalledTimes(3)
    const thirdBody = JSON.parse((fetchMock.mock.calls[2]![1] as { body: string }).body) as Record<
      string,
      unknown
    >
    expect(thirdBody.stream_options).toBeUndefined()
  })

  it('reports usage from a non-streaming JSON body', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        jsonResponse({
          choices: [{ message: { content: 'ok' }, finish_reason: 'stop' }],
          usage: { prompt_tokens: 11, completion_tokens: 4 },
        }),
      ),
    )
    const c = collector()
    await run('openai', c)
    expect(c.usages).toEqual([{ inputTokens: 11, outputTokens: 4 }])
  })
})

describe('anthropic usage', () => {
  it('merges message_start input with message_delta output', async () => {
    vi.stubGlobal(
      'fetch',
      vi
        .fn()
        .mockResolvedValue(
          okResponse(
            sseStream([
              'data: {"type":"message_start","message":{"usage":{"input_tokens":100,"output_tokens":1}}}',
              'data: {"type":"content_block_delta","delta":{"type":"text_delta","text":"hi"}}',
              'data: {"type":"message_delta","delta":{"stop_reason":"end_turn"},"usage":{"output_tokens":42}}',
              '',
            ]),
          ),
        ),
    )
    const c = collector()
    await run('anthropic', c)
    expect(c.deltas.join('')).toBe('hi')
    expect(c.usages).toEqual([{ inputTokens: 100, outputTokens: 42 }])
  })

  it('counts cached prompt tokens as input', async () => {
    vi.stubGlobal(
      'fetch',
      vi
        .fn()
        .mockResolvedValue(
          okResponse(
            sseStream([
              'data: {"type":"message_start","message":{"usage":{"input_tokens":5,"cache_read_input_tokens":900,"cache_creation_input_tokens":95,"output_tokens":1}}}',
              'data: {"type":"content_block_delta","delta":{"type":"text_delta","text":"hi"}}',
              'data: {"type":"message_delta","delta":{"stop_reason":"end_turn"},"usage":{"output_tokens":7}}',
              '',
            ]),
          ),
        ),
    )
    const c = collector()
    await run('anthropic', c)
    expect(c.usages).toEqual([{ inputTokens: 1000, outputTokens: 7 }])
  })
})

describe('gemini usage', () => {
  it('keeps the last chunk usageMetadata', async () => {
    vi.stubGlobal(
      'fetch',
      vi
        .fn()
        .mockResolvedValue(
          okResponse(
            sseStream([
              'data: {"candidates":[{"content":{"parts":[{"text":"hi"}]}}],"usageMetadata":{"promptTokenCount":10,"candidatesTokenCount":2}}',
              'data: {"candidates":[{"finishReason":"STOP"}],"usageMetadata":{"promptTokenCount":12,"candidatesTokenCount":34}}',
              '',
            ]),
          ),
        ),
    )
    const c = collector()
    await run('gemini', c)
    expect(c.deltas.join('')).toBe('hi')
    expect(c.usages).toEqual([{ inputTokens: 12, outputTokens: 34 }])
  })
})
