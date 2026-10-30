import { afterEach, describe, expect, it, vi } from 'vitest'
import type { AgentMessage, AgentToolCall } from '@genoffice/agent-core'
import { chatForProvider } from '../src/chat'
import { streamForProvider } from '../src/stream'
import { jsonResponse, okResponse, sseStream } from './test-utils'

afterEach(() => {
  vi.unstubAllGlobals()
})

function collector() {
  const deltas: string[] = []
  const reasoning: string[] = []
  const toolCalls: AgentToolCall[] = []
  const stopReasons: string[] = []
  const usage: unknown[] = []
  return {
    deltas,
    reasoning,
    toolCalls,
    stopReasons,
    usage,
    cb: {
      signal: new AbortController().signal,
      onDelta: (text: string) => deltas.push(text),
      onReasoningDelta: (text: string) => reasoning.push(text),
      onToolCall: (call: AgentToolCall) => toolCalls.push(call),
      onStopReason: (reason: string) => stopReasons.push(reason),
      onUsage: (u: unknown) => usage.push(u),
    },
  }
}

const ev = (value: unknown) => `data: ${JSON.stringify(value)}`

const REASONING_ITEM = { id: 'rs_1', type: 'reasoning', encrypted_content: 'ENC', summary: [] }
const CALL_ITEM = {
  id: 'fc_1',
  type: 'function_call',
  call_id: 'call_1',
  name: 'get_weather',
  arguments: '{"city":"Paris"}',
}

function toolTurnStream(extra: unknown[] = []) {
  return sseStream([
    ev({ type: 'response.created' }),
    ev({
      type: 'response.output_item.added',
      item: { ...REASONING_ITEM, encrypted_content: undefined },
    }),
    ev({ type: 'response.reasoning_summary_text.delta', delta: 'Need the tool' }),
    ev({ type: 'response.output_item.done', item: REASONING_ITEM }),
    ev({ type: 'response.output_item.added', item: { ...CALL_ITEM, arguments: '' } }),
    ev({ type: 'response.function_call_arguments.delta', item_id: 'fc_1', delta: '{"city":' }),
    ev({ type: 'response.function_call_arguments.delta', item_id: 'fc_1', delta: '"Paris"}' }),
    ev({ type: 'response.output_item.done', item: CALL_ITEM }),
    ...extra.map(ev),
    ev({
      type: 'response.completed',
      response: { status: 'completed', usage: { input_tokens: 63, output_tokens: 48 } },
    }),
  ])
}

const TOOL = {
  name: 'get_weather',
  description: 'Weather for a city',
  inputSchema: { type: 'object', properties: { city: { type: 'string' } } },
}

describe('openai-responses: request shape', () => {
  it('posts to /responses with instructions, input items, flat function tools and no storage', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okResponse(toolTurnStream()))
    vi.stubGlobal('fetch', fetchMock)
    const c = collector()
    await streamForProvider(
      'openai',
      { apiKey: 'k', model: 'gpt-6.1-terra' },
      'be brief',
      [{ role: 'user', text: 'Weather in Paris?' }],
      [TOOL],
      4096,
      c.cb,
    )
    const [url, init] = fetchMock.mock.calls[0] as [string, RequestInit]
    expect(url).toBe('https://api.openai.com/v1/responses')
    const body = JSON.parse(init.body as string)
    expect(body).toMatchObject({
      model: 'gpt-6.1-terra',
      instructions: 'be brief',
      input: [{ role: 'user', content: [{ type: 'input_text', text: 'Weather in Paris?' }] }],
      tools: [
        { type: 'function', name: 'get_weather', parameters: TOOL.inputSchema, strict: false },
      ],
      max_output_tokens: 4096,
      store: false,
      include: ['reasoning.encrypted_content'],
      reasoning: { summary: 'auto' },
      stream: true,
    })
    expect(body.temperature).toBeUndefined()
    expect(body.max_tokens).toBeUndefined()
    expect(body.messages).toBeUndefined()
  })

  it('replays the encrypted reasoning ahead of the function calls, then the tool outputs', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okResponse(toolTurnStream()))
    vi.stubGlobal('fetch', fetchMock)
    const history: AgentMessage[] = [
      { role: 'user', text: 'Weather in Paris?' },
      {
        role: 'assistant',
        text: 'Sure.',
        toolCalls: [
          {
            id: 'call_1',
            name: 'get_weather',
            input: { city: 'Paris' },
            signature: 'oai-rs:{"id":"rs_1","encrypted":"ENC"}',
          },
          { id: 'call_2', name: 'get_weather', input: { city: 'Lyon' } },
        ],
      },
      {
        role: 'tool',
        results: [
          { id: 'call_1', name: 'get_weather', output: 'Sunny' },
          { id: 'call_2', name: 'get_weather', output: 'boom', isError: true },
        ],
      },
      { role: 'assistant', text: '' },
    ]
    await streamForProvider(
      'openai',
      { apiKey: 'k', model: 'gpt-6-astra' },
      'sys',
      history,
      [TOOL],
      100,
      collector().cb,
    )
    const body = JSON.parse((fetchMock.mock.calls[0] as [string, RequestInit])[1].body as string)
    expect(body.input).toEqual([
      { role: 'user', content: [{ type: 'input_text', text: 'Weather in Paris?' }] },
      { type: 'reasoning', id: 'rs_1', encrypted_content: 'ENC', summary: [] },
      { role: 'assistant', content: [{ type: 'output_text', text: 'Sure.' }] },
      {
        type: 'function_call',
        call_id: 'call_1',
        name: 'get_weather',
        arguments: '{"city":"Paris"}',
      },
      {
        type: 'function_call',
        call_id: 'call_2',
        name: 'get_weather',
        arguments: '{"city":"Lyon"}',
      },
      { type: 'function_call_output', call_id: 'call_1', output: 'Sunny' },
      { type: 'function_call_output', call_id: 'call_2', output: 'Error: boom' },
      { role: 'assistant', content: [{ type: 'output_text', text: '(no content)' }] },
    ])
  })

  it('sends images as input_image parts', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okResponse(toolTurnStream()))
    vi.stubGlobal('fetch', fetchMock)
    await streamForProvider(
      'openai',
      { apiKey: 'k', model: 'gpt-5.6-terra' },
      'sys',
      [{ role: 'user', text: 'What is this?', images: [{ mime: 'image/png', base64: 'AAAA' }] }],
      [],
      100,
      collector().cb,
    )
    const body = JSON.parse((fetchMock.mock.calls[0] as [string, RequestInit])[1].body as string)
    expect(body.input[0].content).toEqual([
      { type: 'input_text', text: 'What is this?' },
      { type: 'input_image', image_url: 'data:image/png;base64,AAAA', detail: 'auto' },
    ])
    expect(body.tools).toBeUndefined()
  })
})

describe('openai-responses: stream handling', () => {
  it('streams reasoning summary, emits the finished tool call carrying the reasoning, and reports usage', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(toolTurnStream())))
    const c = collector()
    await streamForProvider(
      'openai',
      { apiKey: 'k', model: 'gpt-6.1-terra' },
      'sys',
      [{ role: 'user', text: 'hi' }],
      [TOOL],
      100,
      c.cb,
    )
    expect(c.reasoning).toEqual(['Need the tool'])
    expect(c.toolCalls).toEqual([
      {
        id: 'call_1',
        name: 'get_weather',
        input: { city: 'Paris' },
        inputError: undefined,
        signature: 'oai-rs:{"id":"rs_1","encrypted":"ENC"}',
      },
    ])
    expect(c.deltas).toEqual([])
    expect(c.stopReasons).toEqual([])
    expect(c.usage).toEqual([{ inputTokens: 63, outputTokens: 48 }])
  })

  it('streams text deltas and does not re-emit the finished message item', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        okResponse(
          sseStream([
            ev({ type: 'response.output_text.delta', delta: 'Hel' }),
            ev({ type: 'response.output_text.delta', delta: 'lo' }),
            ev({
              type: 'response.output_item.done',
              item: { type: 'message', content: [{ type: 'output_text', text: 'Hello' }] },
            }),
            ev({ type: 'response.completed', response: { status: 'completed' } }),
          ]),
        ),
      ),
    )
    const c = collector()
    await streamForProvider(
      'openai',
      { apiKey: 'k', model: 'gpt-6.1-terra' },
      'sys',
      [{ role: 'user', text: 'hi' }],
      [],
      100,
      c.cb,
    )
    expect(c.deltas.join('')).toBe('Hello')
    expect(c.toolCalls).toEqual([])
  })

  it('marks the last tool call truncated and reports max_tokens on an incomplete response', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        okResponse(
          sseStream([
            ev({
              type: 'response.output_item.done',
              item: { ...CALL_ITEM, arguments: '{"city":"Pa' },
            }),
            ev({
              type: 'response.incomplete',
              response: {
                status: 'incomplete',
                incomplete_details: { reason: 'max_output_tokens' },
              },
            }),
          ]),
        ),
      ),
    )
    const c = collector()
    await streamForProvider(
      'openai',
      { apiKey: 'k', model: 'gpt-6.1-terra' },
      'sys',
      [{ role: 'user', text: 'hi' }],
      [TOOL],
      100,
      c.cb,
    )
    expect(c.toolCalls[0]?.truncated).toBe(true)
    expect(c.toolCalls[0]?.inputError).toBeDefined()
    expect(c.stopReasons).toEqual(['max_tokens'])
  })

  it('surfaces response.failed and error events instead of an empty success', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        okResponse(
          sseStream([
            ev({
              type: 'response.failed',
              response: { status: 'failed', error: { message: 'server exploded' } },
            }),
          ]),
        ),
      ),
    )
    await expect(
      streamForProvider(
        'openai',
        { apiKey: 'k', model: 'gpt-6.1-terra' },
        'sys',
        [{ role: 'user', text: 'hi' }],
        [],
        100,
        collector().cb,
      ),
    ).rejects.toThrow(/server exploded/)
  })

  it('a stream that drops mid tool call is reported as a dropped connection', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        okResponse(
          sseStream([
            ev({ type: 'response.output_item.added', item: { ...CALL_ITEM, arguments: '' } }),
            ev({
              type: 'response.function_call_arguments.delta',
              item_id: 'fc_1',
              delta: '{"ci',
            }),
          ]),
        ),
      ),
    )
    await expect(
      streamForProvider(
        'openai',
        { apiKey: 'k', model: 'gpt-6.1-terra' },
        'sys',
        [{ role: 'user', text: 'hi' }],
        [TOOL],
        100,
        collector().cb,
      ),
    ).rejects.toThrow(/connection was dropped/)
  })

  it('an empty completed stream is an error, not a silent success', async () => {
    vi.stubGlobal(
      'fetch',
      vi
        .fn()
        .mockResolvedValue(
          okResponse(
            sseStream([ev({ type: 'response.completed', response: { status: 'completed' } })]),
          ),
        ),
    )
    await expect(
      streamForProvider(
        'openai',
        { apiKey: 'k', model: 'gpt-6.1-terra' },
        'sys',
        [{ role: 'user', text: 'hi' }],
        [],
        100,
        collector().cb,
      ),
    ).rejects.toThrow(/returned no content/)
  })

  it('accepts a complete JSON body in place of SSE', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        jsonResponse({
          status: 'completed',
          output: [
            REASONING_ITEM,
            { type: 'message', content: [{ type: 'output_text', text: 'Sure.' }] },
            CALL_ITEM,
          ],
          usage: { input_tokens: 1, output_tokens: 2 },
        }),
      ),
    )
    const c = collector()
    await streamForProvider(
      'openai',
      { apiKey: 'k', model: 'gpt-6.1-terra' },
      'sys',
      [{ role: 'user', text: 'hi' }],
      [TOOL],
      100,
      c.cb,
    )
    expect(c.deltas).toEqual(['Sure.'])
    expect(c.toolCalls.map((t) => t.name)).toEqual(['get_weather'])
    expect(c.toolCalls[0]?.signature).toContain('rs_1')
    expect(c.usage).toEqual([{ inputTokens: 1, outputTokens: 2 }])
  })

  it('a 4xx surfaces with the body detail', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(new Response('{"error":{"message":"bad model"}}', { status: 400 })),
    )
    await expect(
      streamForProvider(
        'openai',
        { apiKey: 'k', model: 'gpt-6.1-terra' },
        'sys',
        [{ role: 'user', text: 'hi' }],
        [],
        100,
        collector().cb,
      ),
    ).rejects.toThrow(/HTTP 400/)
  })
})

describe('openai-responses: one-shot chat', () => {
  it('reads output_text from the message items of a non-streamed response', async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      jsonResponse({
        status: 'completed',
        output: [
          REASONING_ITEM,
          { type: 'message', content: [{ type: 'output_text', text: 'Hello ' }] },
          { type: 'message', content: [{ type: 'output_text', text: 'there' }] },
        ],
      }),
    )
    vi.stubGlobal('fetch', fetchMock)
    const result = await chatForProvider(
      'openai',
      { apiKey: 'k', model: 'gpt-6.1-terra' },
      'sys',
      'hi',
    )
    expect(result).toEqual({ ok: true, content: 'Hello there' })
    const body = JSON.parse((fetchMock.mock.calls[0] as [string, RequestInit])[1].body as string)
    expect(body.stream).toBe(false)
    expect(body.tools).toBeUndefined()
  })

  it('an empty output is reported as a failed reply', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(jsonResponse({ status: 'completed', output: [] })),
    )
    const result = await chatForProvider(
      'openai',
      { apiKey: 'k', model: 'gpt-6.1-terra' },
      'sys',
      'hi',
    )
    expect(result).toEqual({ ok: false, error: 'AI returned an empty response' })
  })
})
