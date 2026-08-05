import { afterEach, describe, expect, it, vi } from 'vitest'
import type { AgentToolCall } from '@genoffice/agent-core'
import { AiCreditsError, sseLines, streamForProvider } from '../src/stream'
import { jsonResponse, okResponse, sseStream } from './test-utils'

afterEach(() => {
  vi.unstubAllGlobals()
})

function collector() {
  const deltas: string[] = []
  const toolCalls: AgentToolCall[] = []
  const stopReasons: string[] = []
  return {
    deltas,
    toolCalls,
    stopReasons,
    cb: {
      signal: new AbortController().signal,
      onDelta: (text: string) => deltas.push(text),
      onToolCall: (call: AgentToolCall) => toolCalls.push(call),
      onStopReason: (reason: string) => stopReasons.push(reason),
    },
  }
}

describe('sseLines', () => {
  it('splits a stream into lines, including a trailing line with no newline', async () => {
    const encoder = new TextEncoder()
    const body = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(encoder.encode('data: a\ndata: b\n'))
        controller.enqueue(encoder.encode('data: c')) // no trailing newline
        controller.close()
      },
    })
    const lines: string[] = []
    for await (const line of sseLines(body)) lines.push(line)
    expect(lines).toEqual(['data: a', 'data: b', 'data: c'])
  })
})

describe('streamForProvider: anthropic', () => {
  it('emits text deltas and a completed tool call', async () => {
    const body = sseStream([
      'data: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"hello "}}',
      'data: {"type":"content_block_start","index":1,"content_block":{"type":"tool_use","id":"t1","name":"do_thing"}}',
      'data: {"type":"content_block_delta","index":1,"delta":{"type":"input_json_delta","partial_json":"{\\"a\\":"}}',
      'data: {"type":"content_block_delta","index":1,"delta":{"type":"input_json_delta","partial_json":"1}"}}',
      'data: {"type":"content_block_stop","index":1}',
      'data: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"world"}}',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { deltas, toolCalls, cb } = collector()
    await streamForProvider(
      'anthropic',
      { apiKey: 'k', model: 'claude-sonnet-5' },
      'system',
      [{ role: 'user', text: 'hi' }],
      [],
      100,
      cb,
    )
    expect(deltas.join('')).toBe('hello world')
    expect(toolCalls).toEqual([{ id: 't1', name: 'do_thing', input: { a: 1 } }])
  })

  it('repairs unescaped quotes inside tool input string values', async () => {
    const partial = JSON.stringify({
      type: 'content_block_delta',
      index: 1,
      delta: { type: 'input_json_delta', partial_json: '{"topic": "from "future" to "present""}' },
    })
    const body = sseStream([
      'data: {"type":"content_block_start","index":1,"content_block":{"type":"tool_use","id":"t1","name":"gen"}}',
      `data: ${partial}`,
      'data: {"type":"content_block_stop","index":1}',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { toolCalls, cb } = collector()
    await streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(toolCalls).toEqual([
      { id: 't1', name: 'gen', input: { topic: 'from "future" to "present"' } },
    ])
  })

  it('unparseable tool input becomes inputError instead of killing the stream', async () => {
    const partial = JSON.stringify({
      type: 'content_block_delta',
      index: 1,
      delta: { type: 'input_json_delta', partial_json: '{"a": 1,' }, // truncated JSON, unrepairable
    })
    const body = sseStream([
      'data: {"type":"content_block_start","index":1,"content_block":{"type":"tool_use","id":"t1","name":"gen"}}',
      `data: ${partial}`,
      'data: {"type":"content_block_stop","index":1}',
      'data: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"after"}}',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { deltas, toolCalls, cb } = collector()
    await streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(toolCalls).toHaveLength(1)
    expect(toolCalls[0]!.input).toEqual({})
    expect(toolCalls[0]!.inputError).toContain('raw: {"a": 1,')
    expect(deltas.join('')).toBe('after') // the stream was not interrupted
  })

  it('surfaces message_delta stop_reason and does not flag complete tool calls', async () => {
    const body = sseStream([
      'data: {"type":"content_block_start","index":1,"content_block":{"type":"tool_use","id":"t1","name":"gen"}}',
      'data: {"type":"content_block_delta","index":1,"delta":{"type":"input_json_delta","partial_json":"{\\"a\\":1}"}}',
      'data: {"type":"content_block_stop","index":1}',
      'data: {"type":"message_delta","delta":{"stop_reason":"end_turn"}}',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { toolCalls, stopReasons, cb } = collector()
    await streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(stopReasons).toEqual(['end_turn'])
    expect(toolCalls[0]!.truncated).toBeUndefined()
  })

  it('max_tokens marks the cut-off tool call as truncated and reports the stop reason', async () => {
    const body = sseStream([
      'data: {"type":"content_block_start","index":1,"content_block":{"type":"tool_use","id":"t1","name":"gen"}}',
      'data: {"type":"content_block_delta","index":1,"delta":{"type":"input_json_delta","partial_json":"{\\"html\\": \\"<p>very lo"}}',
      'data: {"type":"content_block_stop","index":1}',
      'data: {"type":"message_delta","delta":{"stop_reason":"max_tokens"}}',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { toolCalls, stopReasons, cb } = collector()
    await streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(stopReasons).toEqual(['max_tokens'])
    expect(toolCalls).toHaveLength(1)
    expect(toolCalls[0]!.truncated).toBe(true)
    expect(toolCalls[0]!.inputError).toBeDefined()
  })

  it('throws on a non-ok HTTP response', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response('bad key', { status: 401 })))
    const { cb } = collector()
    await expect(
      streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toThrow(/Claude HTTP 401/)
  })

  it('throws on an Anthropic-protocol error event', async () => {
    const body = sseStream(['data: {"type":"error","error":{"message":"Overloaded"}}'])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { cb } = collector()
    await expect(
      streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toThrow('Overloaded')
  })

  it('throws on a gateway error event that does not follow the Anthropic protocol', async () => {
    // e.g. a proxy reporting quota exhaustion in OpenAI shape on the Anthropic route:
    // previously this was silently ignored and surfaced as an empty "successful" turn
    const body = sseStream(['data: {"error":{"message":"Insufficient credits"}}'])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { cb } = collector()
    await expect(
      streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toThrow('Insufficient credits')
  })

  it('emits the text of a complete JSON message sent instead of an SSE stream', async () => {
    // Gateways can answer stream:true with a complete JSON message; it must not
    // dissolve into an empty "successful" turn (credits notices throw instead — see below)
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        jsonResponse({
          type: 'message',
          role: 'assistant',
          content: [{ type: 'text', text: 'Scheduled maintenance at 06:00 UTC.' }],
          stop_reason: 'end_turn',
        }),
      ),
    )
    const { deltas, stopReasons, cb } = collector()
    await streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(deltas.join('')).toBe('Scheduled maintenance at 06:00 UTC.')
    expect(stopReasons).toEqual(['end_turn'])
  })

  it('emits tool calls from a complete JSON message sent instead of an SSE stream', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        jsonResponse({
          content: [{ type: 'tool_use', id: 't1', name: 'do_thing', input: { a: 1 } }],
        }),
      ),
    )
    const { toolCalls, cb } = collector()
    await streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(toolCalls).toEqual([{ id: 't1', name: 'do_thing', input: { a: 1 } }])
  })

  it('flags the last tool call of a max_tokens JSON body as truncated', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        jsonResponse({
          content: [{ type: 'tool_use', id: 't1', name: 'do_thing', input: { a: 1 } }],
          stop_reason: 'max_tokens',
        }),
      ),
    )
    const { toolCalls, stopReasons, cb } = collector()
    await streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(toolCalls[0]!.truncated).toBe(true)
    expect(stopReasons).toEqual(['max_tokens'])
  })

  it('throws on an empty or error-bearing JSON body sent instead of an SSE stream', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(jsonResponse({ error: { message: 'Insufficient credits' } })),
    )
    const { cb } = collector()
    await expect(
      streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toThrow('Insufficient credits')

    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(jsonResponse({ content: [] })))
    await expect(
      streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toThrow(/Claude returned no content/)
  })

  it('never sends an empty assistant content array when history has edits-only replies', async () => {
    // Prior empty terminal turns would map to content:[] and break follow-ups
    // on Anthropic (genoffice#12 / #22 class of multi-turn failures).
    const fetchMock = vi
      .fn()
      .mockResolvedValue(
        okResponse(
          sseStream([
            'data: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"ok"}}',
            'data: {"type":"message_delta","delta":{"stop_reason":"end_turn"}}',
          ]),
        ),
      )
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider(
      'anthropic',
      { apiKey: 'k', model: 'm' },
      'sys',
      [
        { role: 'user', text: 'first' },
        { role: 'assistant', text: '' },
        { role: 'user', text: 'second' },
      ],
      [],
      100,
      cb,
    )
    const body = JSON.parse(fetchMock.mock.calls[0]![1].body as string) as {
      messages: Array<{ role: string; content: unknown }>
    }
    for (const msg of body.messages) {
      if (msg.role !== 'assistant') continue
      expect(Array.isArray(msg.content)).toBe(true)
      expect((msg.content as unknown[]).length).toBeGreaterThan(0)
    }
  })

  it('replaces an HTML error body (e.g. a gateway block page) with a readable note', async () => {
    const html =
      '<!doctype html>\n<html>\n<head><title>Genspark</title></head><body>app shell</body></html>'
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response(html, { status: 403 })))
    const { cb } = collector()
    await expect(
      streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toThrow(/Claude HTTP 403: .*web page instead of an API response/)
  })
})

describe('streamForProvider: gemini', () => {
  it('emits text and a whole (non-partial) function call', async () => {
    const body = sseStream([
      'data: {"candidates":[{"content":{"parts":[{"text":"hi there"}]}}]}',
      'data: {"candidates":[{"content":{"parts":[{"functionCall":{"name":"set_cell","args":{"a1":"42"}}}]}}]}',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { deltas, toolCalls, cb } = collector()
    await streamForProvider(
      'gemini',
      { apiKey: 'k', model: 'gemini-2.5-flash' },
      'sys',
      [],
      [],
      100,
      cb,
    )
    expect(deltas.join('')).toBe('hi there')
    expect(toolCalls).toHaveLength(1)
    expect(toolCalls[0]).toMatchObject({ name: 'set_cell', input: { a1: '42' } })
  })

  it('throws when the prompt is blocked instead of finishing an empty turn', async () => {
    const body = sseStream(['data: {"promptFeedback":{"blockReason":"SAFETY"}}'])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { cb } = collector()
    await expect(
      streamForProvider('gemini', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toThrow(/blocked the prompt \(SAFETY\)/)
  })

  it('throws on an abnormal finishReason that produced no content', async () => {
    const body = sseStream(['data: {"candidates":[{"finishReason":"RECITATION"}]}'])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { cb } = collector()
    await expect(
      streamForProvider('gemini', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toThrow(/no content \(finishReason=RECITATION\)/)
  })

  it('keeps partial content when an abnormal finishReason arrives after text', async () => {
    const body = sseStream([
      'data: {"candidates":[{"content":{"parts":[{"text":"partial answer"}]}}]}',
      'data: {"candidates":[{"finishReason":"SAFETY"}]}',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { deltas, cb } = collector()
    await streamForProvider('gemini', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(deltas.join('')).toBe('partial answer')
  })

  it('throws on a gateway error event', async () => {
    const body = sseStream(['data: {"error":{"message":"quota exceeded"}}'])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { cb } = collector()
    await expect(
      streamForProvider('gemini', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toThrow('quota exceeded')
  })

  it('emits content from a complete JSON body (object or chunk array) sent instead of SSE', async () => {
    vi.stubGlobal(
      'fetch',
      vi
        .fn()
        .mockResolvedValue(
          jsonResponse([
            { candidates: [{ content: { parts: [{ text: 'chunk one ' }] } }] },
            { candidates: [{ content: { parts: [{ text: 'chunk two' }] }, finishReason: 'STOP' }] },
          ]),
        ),
    )
    const { deltas, cb } = collector()
    await streamForProvider('gemini', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(deltas.join('')).toBe('chunk one chunk two')
  })

  it('throws on an empty JSON body sent instead of SSE', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(jsonResponse({ candidates: [] })))
    const { cb } = collector()
    await expect(
      streamForProvider('gemini', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toThrow(/Gemini returned no content/)
  })

  it('surfaces MAX_TOKENS and abnormal finishReason from a JSON body sent instead of SSE', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        jsonResponse({
          candidates: [{ content: { parts: [{ text: 'cut off' }] }, finishReason: 'MAX_TOKENS' }],
        }),
      ),
    )
    const { deltas, stopReasons, cb } = collector()
    await streamForProvider('gemini', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(deltas.join('')).toBe('cut off')
    expect(stopReasons).toEqual(['max_tokens'])

    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(jsonResponse({ candidates: [{ finishReason: 'SAFETY' }] })),
    )
    await expect(
      streamForProvider('gemini', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toThrow(/no content \(finishReason=SAFETY\)/)
  })
})

describe('streamForProvider: openai-compatible', () => {
  it('reassembles fragmented tool call arguments and flushes on finish_reason', async () => {
    const body = sseStream([
      'data: {"choices":[{"delta":{"content":"partial "}}]}',
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"c1","function":{"name":"replace"}}]}}]}',
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":"{\\"x\\":1}"}}]}}]}',
      'data: {"choices":[{"delta":{},"finish_reason":"tool_calls"}]}',
      'data: [DONE]',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { deltas, toolCalls, cb } = collector()
    await streamForProvider(
      'openai',
      { apiKey: 'k', model: 'gpt-4.1-mini' },
      'sys',
      [],
      [],
      100,
      cb,
    )
    expect(deltas.join('')).toBe('partial ')
    expect(toolCalls).toEqual([{ id: 'c1', name: 'replace', input: { x: 1 } }])
  })

  it("finish_reason 'length' normalizes to max_tokens and flags the cut-off tool call", async () => {
    const body = sseStream([
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"c1","function":{"name":"replace","arguments":"{\\"x\\": \\"trunc"}}]}}]}',
      'data: {"choices":[{"delta":{},"finish_reason":"length"}]}',
      'data: [DONE]',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { toolCalls, stopReasons, cb } = collector()
    await streamForProvider('openai', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(stopReasons).toEqual(['max_tokens'])
    expect(toolCalls).toHaveLength(1)
    expect(toolCalls[0]!.truncated).toBe(true)
    expect(toolCalls[0]!.inputError).toBeDefined()
  })

  it('throws on a gateway error event instead of finishing an empty turn', async () => {
    const body = sseStream([
      'data: {"error":{"message":"You exceeded your current quota"}}',
      'data: [DONE]',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { cb } = collector()
    await expect(
      streamForProvider('openai', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toThrow('You exceeded your current quota')
  })

  it('throws when a content_filter finish produced no content', async () => {
    const body = sseStream([
      'data: {"choices":[{"delta":{},"finish_reason":"content_filter"}]}',
      'data: [DONE]',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { cb } = collector()
    await expect(
      streamForProvider('openai', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toThrow(/no content \(finish_reason=content_filter\)/)
  })

  it('keeps partial content when content_filter cuts off after some text', async () => {
    const body = sseStream([
      'data: {"choices":[{"delta":{"content":"partial "}}]}',
      'data: {"choices":[{"delta":{},"finish_reason":"content_filter"}]}',
      'data: [DONE]',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { deltas, cb } = collector()
    await streamForProvider('openai', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(deltas.join('')).toBe('partial ')
  })

  it('emits content and tool calls from a complete JSON body sent instead of SSE', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        jsonResponse({
          choices: [
            {
              message: {
                content: 'Here is the change.',
                tool_calls: [{ id: 'c1', function: { name: 'replace', arguments: '{"x":1}' } }],
              },
              finish_reason: 'stop',
            },
          ],
        }),
      ),
    )
    const { deltas, toolCalls, cb } = collector()
    await streamForProvider('openai', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(deltas.join('')).toBe('Here is the change.')
    expect(toolCalls).toEqual([
      { id: 'c1', name: 'replace', input: { x: 1 }, inputError: undefined },
    ])
  })

  it("flags the last tool call of a JSON body with finish_reason 'length' as truncated", async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        jsonResponse({
          choices: [
            {
              message: {
                tool_calls: [{ id: 'c1', function: { name: 'replace', arguments: '{"x": "tru' } }],
              },
              finish_reason: 'length',
            },
          ],
        }),
      ),
    )
    const { toolCalls, stopReasons, cb } = collector()
    await streamForProvider('openai', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(toolCalls[0]!.truncated).toBe(true)
    expect(toolCalls[0]!.inputError).toBeDefined()
    expect(stopReasons).toEqual(['max_tokens'])
  })

  it('throws on an empty JSON body sent instead of SSE', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(jsonResponse({ choices: [{ message: {} }] })))
    const { cb } = collector()
    await expect(
      streamForProvider('openai', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toThrow(/The model returned no content/)
  })

  it('never sends content:null for an assistant turn with neither text nor tools', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(
        okResponse(
          sseStream([
            'data: {"choices":[{"delta":{"content":"ok"}}]}',
            'data: {"choices":[{"delta":{},"finish_reason":"stop"}]}',
            'data: [DONE]',
          ]),
        ),
      )
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider(
      'openai',
      { apiKey: 'k', model: 'm' },
      'sys',
      [
        { role: 'user', text: 'first' },
        { role: 'assistant', text: '' },
        { role: 'user', text: 'second' },
      ],
      [],
      100,
      cb,
    )
    const body = JSON.parse(fetchMock.mock.calls[0]![1].body as string) as {
      messages: Array<{ role: string; content: unknown; tool_calls?: unknown }>
    }
    for (const msg of body.messages) {
      if (msg.role !== 'assistant') continue
      expect(msg.content === null && !msg.tool_calls).toBe(false)
      if (msg.content !== null) expect(String(msg.content).length).toBeGreaterThan(0)
    }
  })

  it('routes deepseek and openai to their fixed base URLs', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okResponse(sseStream(['data: [DONE]'])))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider(
      'deepseek',
      { apiKey: 'k', model: 'deepseek-chat' },
      'sys',
      [],
      [],
      100,
      cb,
    )
    expect(fetchMock).toHaveBeenCalledWith(
      'https://api.deepseek.com/v1/chat/completions',
      expect.anything(),
    )
  })

  it('uses the configured base URL for the custom provider', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okResponse(sseStream(['data: [DONE]'])))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider(
      'custom',
      { apiKey: 'k', model: 'm', baseUrl: 'https://my-endpoint.example.com/v1' },
      'sys',
      [],
      [],
      100,
      cb,
    )
    expect(fetchMock).toHaveBeenCalledWith(
      'https://my-endpoint.example.com/v1/chat/completions',
      expect.anything(),
    )
  })

  it('rejects the custom provider without a base URL, without ever calling fetch', async () => {
    const fetchMock = vi.fn()
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await expect(
      streamForProvider('custom', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toThrow(/Base URL/)
    expect(fetchMock).not.toHaveBeenCalled()
  })
})

describe('streamForProvider: genspark', () => {
  it('routes claude models to the Anthropic-compatible proxy endpoint', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okResponse(sseStream([])))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider(
      'genspark',
      { apiKey: 'gsk-k', model: 'claude-opus-4-7' },
      'sys',
      [],
      [],
      100,
      cb,
    )
    expect(fetchMock).toHaveBeenCalledWith(
      'https://www.genspark.ai/api/anthropic/v1/messages',
      expect.objectContaining({ headers: expect.objectContaining({ 'x-api-key': 'gsk-k' }) }),
    )
  })

  it('routes gemini models to the Gemini proxy with header auth', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okResponse(sseStream([])))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider(
      'genspark',
      { apiKey: 'gsk-k', model: 'gemini-3-flash-preview' },
      'sys',
      [],
      [],
      100,
      cb,
    )
    expect(fetchMock).toHaveBeenCalledWith(
      'https://www.genspark.ai/api/llm_proxy/gemini/v1beta/models/gemini-3-flash-preview:streamGenerateContent?alt=sse',
      expect.objectContaining({ headers: expect.objectContaining({ 'x-goog-api-key': 'gsk-k' }) }),
    )
  })

  it('routes other models to the OpenAI-compatible proxy', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okResponse(sseStream(['data: [DONE]'])))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider(
      'genspark',
      { apiKey: 'gsk-k', model: 'gpt-5.2' },
      'sys',
      [],
      [],
      100,
      cb,
    )
    expect(fetchMock).toHaveBeenCalledWith(
      'https://www.genspark.ai/api/llm_proxy/v1/chat/completions',
      expect.anything(),
    )
  })

  it('stamps X-Agent-Type on all three proxy routes for billing attribution', async () => {
    for (const model of ['claude-opus-4-7', 'gemini-3-flash-preview', 'gpt-5.2']) {
      const fetchMock = vi.fn().mockResolvedValue(okResponse(sseStream([])))
      vi.stubGlobal('fetch', fetchMock)
      const { cb } = collector()
      await streamForProvider('genspark', { apiKey: 'gsk-k', model }, 'sys', [], [], 100, cb)
      expect(fetchMock).toHaveBeenCalledWith(
        expect.anything(),
        expect.objectContaining({
          headers: expect.objectContaining({ 'X-Agent-Type': 'genoffice' }),
        }),
      )
    }
  })

  it('never sends X-Agent-Type to direct vendor APIs', async () => {
    for (const [provider, model] of [
      ['anthropic', 'claude-opus-4-7'],
      ['gemini', 'gemini-2.5-flash'],
      ['openai', 'gpt-4.1-mini'],
    ] as const) {
      const fetchMock = vi.fn().mockResolvedValue(okResponse(sseStream([])))
      vi.stubGlobal('fetch', fetchMock)
      const { cb } = collector()
      await streamForProvider(provider, { apiKey: 'k', model }, 'sys', [], [], 100, cb)
      const headers = fetchMock.mock.calls[0]![1].headers as Record<string, string>
      expect(headers['X-Agent-Type']).toBeUndefined()
    }
  })
})

describe('streamForProvider: 200 + non-stream JSON instead of SSE', () => {
  const creditsNotice =
    'Your Genspark credits have been exhausted. Please visit https://www.genspark.ai/pricing to purchase more credits.'
  const json = (value: unknown) =>
    new Response(JSON.stringify(value), {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
    })

  it('anthropic route: a credits-exhausted notice becomes AiCreditsError with the notice text', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        json({
          type: 'message',
          role: 'assistant',
          content: [{ type: 'text', text: creditsNotice }],
          stop_reason: 'end_turn',
          usage: { input_tokens: 0, output_tokens: 0 },
        }),
      ),
    )
    const { deltas, cb } = collector()
    const run = streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    await expect(run).rejects.toBeInstanceOf(AiCreditsError)
    await expect(run).rejects.toThrow(/credits have been exhausted/)
    expect(deltas).toEqual([])
  })

  it('gemini route: zero usage + a pricing link counts as a credits notice', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        json({
          candidates: [
            {
              content: {
                parts: [{ text: 'Out of quota, visit https://www.genspark.ai/pricing to top up.' }],
              },
            },
          ],
          usageMetadata: { promptTokenCount: 0, candidatesTokenCount: 0, totalTokenCount: 0 },
        }),
      ),
    )
    const { cb } = collector()
    await expect(
      streamForProvider('gemini', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toBeInstanceOf(AiCreditsError)
  })

  it('openai route: an insufficient-credits message becomes AiCreditsError', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        json({
          choices: [{ message: { role: 'assistant', content: 'Insufficient credits remaining.' } }],
          usage: { prompt_tokens: 0, completion_tokens: 0 },
        }),
      ),
    )
    const { cb } = collector()
    await expect(
      streamForProvider('openai', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toBeInstanceOf(AiCreditsError)
  })

  it('a non-credits notice is emitted as the reply text instead of an empty turn', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        json({
          type: 'message',
          content: [{ type: 'text', text: 'The service is under maintenance until 06:00 UTC.' }],
          usage: { input_tokens: 3, output_tokens: 12 },
        }),
      ),
    )
    const { deltas, cb } = collector()
    await streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(deltas.join('')).toBe('The service is under maintenance until 06:00 UTC.')
  })

  it('an unextractable body throws with a body summary', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        new Response('not json at all', {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
        }),
      ),
    )
    const { cb } = collector()
    await expect(
      streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toThrow(/Claude returned an unparseable JSON body: not json at all/)
  })

  it('JSON without a message text also falls back to the body summary', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(json({ choices: [] })))
    const { cb } = collector()
    await expect(
      streamForProvider('openai', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toThrow(/The model returned no content: \{"choices":\[\]\}/)
  })

  it('a missing Content-Type is still treated as a stream', async () => {
    const body = sseStream([
      'data: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"ok"}}',
    ])
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(body)))
    const { deltas, cb } = collector()
    await streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb)
    expect(deltas.join('')).toBe('ok')
  })
})

it('rejects an unknown provider id', async () => {
  const { cb } = collector()
  await expect(
    streamForProvider('unknown' as never, { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
  ).rejects.toThrow(/Unknown provider/)
})
