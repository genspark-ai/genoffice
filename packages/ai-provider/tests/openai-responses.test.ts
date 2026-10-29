import { afterEach, describe, expect, it, vi } from 'vitest'
import { streamForProvider } from '../src/stream'
import { getProviderAdapter } from '../src/registry'
import { jsonResponse, okResponse, sseStream } from './test-utils'

afterEach(() => {
  vi.unstubAllGlobals()
})

const TOOL = [{ name: 'read_file', description: 'read a file', inputSchema: { type: 'object' } }]

const TOOL_CALL_SSE = () =>
  sseStream([
    'data: {"type":"response.output_item.added","item":{"type":"function_call","id":"fc_1","call_id":"call_1","name":"read_file","arguments":""}}',
    'data: {"type":"response.function_call_arguments.delta","item_id":"fc_1","delta":"{\\"path\\":"}',
    'data: {"type":"response.function_call_arguments.delta","item_id":"fc_1","delta":"\\"a.txt\\"}"}',
    'data: {"type":"response.output_item.done","item":{"type":"function_call","id":"fc_1","call_id":"call_1","name":"read_file","arguments":"{\\"path\\":\\"a.txt\\"}"}}',
    'data: {"type":"response.completed","response":{"status":"completed","usage":{"input_tokens":11,"output_tokens":7}}}',
  ])

const TEXT_SSE = () =>
  sseStream([
    'data: {"type":"response.output_text.delta","delta":"Hello"}',
    'data: {"type":"response.reasoning_summary_text.delta","delta":"thinking"}',
    'data: {"type":"response.output_text.delta","delta":" world"}',
    'data: {"type":"response.completed","response":{"status":"completed","usage":{"input_tokens":5,"output_tokens":3}}}',
  ])

function lastCallBody(): Record<string, unknown> {
  const body = vi.mocked(fetch).mock.calls.at(-1)?.[1]?.body
  return JSON.parse(String(body)) as Record<string, unknown>
}

describe('openai-responses protocol', () => {
  it('routes the gpt-5.6 family to the Responses API and gpt-5.5 to chat/completions', () => {
    const adapter = getProviderAdapter('openai')
    expect(adapter.resolveEndpoint({ apiKey: 'k', model: 'gpt-5.6-terra' }).protocol).toBe(
      'openai-responses',
    )
    expect(adapter.resolveEndpoint({ apiKey: 'k', model: 'gpt-5.6-luna' }).protocol).toBe(
      'openai-responses',
    )
    expect(adapter.resolveEndpoint({ apiKey: 'k', model: 'gpt-5.5' }).protocol).toBe(
      'openai-compatible',
    )
    // other vendors never move: their gateways do not serve /v1/responses
    expect(
      getProviderAdapter('openrouter').resolveEndpoint({
        apiKey: 'k',
        model: 'openai/gpt-5.6-terra',
      }).protocol,
    ).toBe('openai-compatible')
  })

  it('a gpt-5.6 tool turn posts typed input items and flat tools to /responses, keeping reasoning', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => okResponse(TOOL_CALL_SSE())),
    )
    const toolCalls: unknown[] = []
    const usages: unknown[] = []
    await streamForProvider(
      'openai',
      { apiKey: 'k', model: 'gpt-5.6-terra' },
      'edit the deck',
      [{ role: 'user', text: 'rename slide 2' }],
      TOOL,
      4096,
      {
        signal: new AbortController().signal,
        onDelta: () => {},
        onToolCall: (c) => toolCalls.push(c),
        onUsage: (u) => usages.push(u),
      },
    )
    const body = lastCallBody()
    expect(body.model).toBe('gpt-5.6-terra')
    expect(body.instructions).toBe('edit the deck')
    expect(body.store).toBe(false)
    expect(body.stream).toBe(true)
    expect(body.max_output_tokens).toBe(4096)
    // no temperature on the Responses path, no reasoning_effort pin: the model
    // keeps its default reasoning beside tools
    expect(body.temperature).toBeUndefined()
    expect(body.reasoning).toBeUndefined()
    expect(body.input).toEqual([{ role: 'user', content: 'rename slide 2' }])
    expect(body.tools).toEqual([
      {
        type: 'function',
        name: 'read_file',
        description: 'read a file',
        parameters: { type: 'object' },
        strict: false,
      },
    ])
    expect(toolCalls).toEqual([
      { id: 'call_1', name: 'read_file', input: { path: 'a.txt' }, inputError: undefined },
    ])
    expect(usages).toEqual([{ inputTokens: 11, outputTokens: 7 }])
  })

  it('streams text and reasoning deltas', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => okResponse(TEXT_SSE())),
    )
    const deltas: string[] = []
    const reasoning: string[] = []
    await streamForProvider(
      'openai',
      { apiKey: 'k', model: 'gpt-5.6-terra' },
      'sys',
      [{ role: 'user', text: 'hi' }],
      [],
      4096,
      {
        signal: new AbortController().signal,
        onDelta: (t) => deltas.push(t),
        onToolCall: () => {},
        onReasoningDelta: (t) => reasoning.push(t),
      },
    )
    expect(deltas.join('')).toBe('Hello world')
    expect(reasoning.join('')).toBe('thinking')
  })

  it('replays a tool loop as function_call plus function_call_output items', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => okResponse(TEXT_SSE())),
    )
    await streamForProvider(
      'openai',
      { apiKey: 'k', model: 'gpt-5.6-terra' },
      'sys',
      [
        { role: 'user', text: 'do it' },
        {
          role: 'assistant',
          text: 'reading',
          toolCalls: [{ id: 'call_1', name: 'read_file', input: { path: 'a.txt' } }],
        },
        { role: 'tool', results: [{ id: 'call_1', name: 'read_file', output: 'file body' }] },
      ],
      TOOL,
      4096,
      { signal: new AbortController().signal, onDelta: () => {}, onToolCall: () => {} },
    )
    expect(lastCallBody().input).toEqual([
      { role: 'user', content: 'do it' },
      { role: 'assistant', content: [{ type: 'output_text', text: 'reading' }] },
      {
        type: 'function_call',
        call_id: 'call_1',
        name: 'read_file',
        arguments: '{"path":"a.txt"}',
      },
      { type: 'function_call_output', call_id: 'call_1', output: 'file body' },
    ])
  })

  it('user images travel as input_image parts', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => okResponse(TEXT_SSE())),
    )
    await streamForProvider(
      'openai',
      { apiKey: 'k', model: 'gpt-5.6-terra' },
      'sys',
      [{ role: 'user', text: 'look', images: [{ base64: 'QUJD', mime: 'image/png' }] }],
      [],
      4096,
      { signal: new AbortController().signal, onDelta: () => {}, onToolCall: () => {} },
    )
    expect(lastCallBody().input).toEqual([
      {
        role: 'user',
        content: [
          { type: 'input_text', text: 'look' },
          { type: 'input_image', image_url: 'data:image/png;base64,QUJD' },
        ],
      },
    ])
  })

  it('an incomplete turn reports max_tokens and carries usage', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () =>
        okResponse(
          sseStream([
            'data: {"type":"response.output_text.delta","delta":"partial"}',
            'data: {"type":"response.incomplete","response":{"status":"incomplete","incomplete_details":{"reason":"max_output_tokens"},"usage":{"input_tokens":9,"output_tokens":4096}}}',
          ]),
        ),
      ),
    )
    const stops: string[] = []
    const usages: unknown[] = []
    await streamForProvider(
      'openai',
      { apiKey: 'k', model: 'gpt-5.6-terra' },
      'sys',
      [{ role: 'user', text: 'hi' }],
      [],
      4096,
      {
        signal: new AbortController().signal,
        onDelta: () => {},
        onToolCall: () => {},
        onStopReason: (r) => stops.push(r),
        onUsage: (u) => usages.push(u),
      },
    )
    expect(stops).toEqual(['max_tokens'])
    expect(usages).toEqual([{ inputTokens: 9, outputTokens: 4096 }])
  })

  it('a failed response surfaces the API error message', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () =>
        okResponse(
          sseStream([
            'data: {"type":"response.failed","response":{"error":{"message":"boom from the API"}}}',
          ]),
        ),
      ),
    )
    await expect(
      streamForProvider(
        'openai',
        { apiKey: 'k', model: 'gpt-5.6-terra' },
        'sys',
        [{ role: 'user', text: 'hi' }],
        [],
        4096,
        { signal: new AbortController().signal, onDelta: () => {}, onToolCall: () => {} },
      ),
    ).rejects.toThrow('boom from the API')
  })

  it('a whole JSON body swallowed by a gateway still yields text and tool calls', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () =>
        jsonResponse({
          status: 'completed',
          output: [{ type: 'message', content: [{ type: 'output_text', text: 'plain answer' }] }],
          usage: { input_tokens: 3, output_tokens: 4 },
        }),
      ),
    )
    const deltas: string[] = []
    await streamForProvider(
      'openai',
      { apiKey: 'k', model: 'gpt-5.6-terra' },
      'sys',
      [{ role: 'user', text: 'hi' }],
      [],
      4096,
      {
        signal: new AbortController().signal,
        onDelta: (t) => deltas.push(t),
        onToolCall: () => {},
      },
    )
    expect(deltas.join('')).toBe('plain answer')
  })
})
