import { afterEach, describe, expect, it, vi } from 'vitest'
import type { AgentToolCall } from '@genoffice/agent-core'
import { sanitizeGeminiSchema, sseLines, streamForProvider } from '../src/stream'
import { okResponse, sseStream } from './test-utils'

afterEach(() => {
  vi.unstubAllGlobals()
})

function collector() {
  const deltas: string[] = []
  const toolCalls: AgentToolCall[] = []
  return {
    deltas,
    toolCalls,
    cb: {
      signal: new AbortController().signal,
      onDelta: (text: string) => deltas.push(text),
      onToolCall: (call: AgentToolCall) => toolCalls.push(call),
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

  it('throws on a non-ok HTTP response', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response('bad key', { status: 401 })))
    const { cb } = collector()
    await expect(
      streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toThrow(/Claude HTTP 401/)
  })

  it('replaces an HTML error body (e.g. a gateway block page) with a readable note', async () => {
    const html = '<!doctype html>\n<html>\n<head><title>Genspark</title></head><body>app shell</body></html>'
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response(html, { status: 403 })))
    const { cb } = collector()
    await expect(
      streamForProvider('anthropic', { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
    ).rejects.toThrow(/Claude HTTP 403: .*web page instead of an API response/)
  })
})

describe('sanitizeGeminiSchema', () => {
  it('collapses type unions to a scalar and marks nullable', () => {
    expect(sanitizeGeminiSchema({ type: ['number', 'null'], description: 'v' })).toEqual({
      type: 'number',
      description: 'v',
      nullable: true,
    })
    expect(sanitizeGeminiSchema({ type: ['string', 'number'] })).toEqual({
      type: 'string',
    })
    expect(sanitizeGeminiSchema({ type: ['null'] })).toEqual({ type: 'string', nullable: true })
  })

  it('normalizes unions nested inside array items', () => {
    const out = sanitizeGeminiSchema({
      type: 'object',
      properties: {
        values: { type: 'array', items: { type: ['number', 'null'] } },
      },
    }) as { properties: { values: { items: { type: string; nullable: boolean } } } }
    expect(out.properties.values.items).toEqual({ type: 'number', nullable: true })
  })

  it('strips additionalProperties, $schema, propertyNames and definitions', () => {
    const out = sanitizeGeminiSchema({
      type: 'object',
      $schema: 'http://json-schema.org/draft-07/schema#',
      propertyNames: { type: 'string' },
      additionalProperties: false,
      definitions: { foo: { type: 'string' } },
      properties: { tags: { type: 'array', items: { type: 'string' }, additionalProperties: false } },
    }) as Record<string, unknown>
    expect(out.$schema).toBeUndefined()
    expect(out.propertyNames).toBeUndefined()
    expect(out.additionalProperties).toBeUndefined()
    expect(out.definitions).toBeUndefined()
    expect((out.properties as Record<string, unknown>).tags).not.toHaveProperty('additionalProperties')
  })

  it('collapses anyOf/oneOf/allOf to the first non-null member', () => {
    const out = sanitizeGeminiSchema({
      type: 'object',
      properties: {
        x: { anyOf: [{ type: 'null' }, { type: 'string', description: 'kept' }] },
      },
    }) as { properties: { x: { type: string; description?: string } } }
    expect(out.properties.x.type).toBe('string')
    expect(out.properties.x.description).toBe('kept')
  })

  it('keeps empty properties objects (accepted by Gemini) intact', () => {
    expect(sanitizeGeminiSchema({ type: 'object', properties: {}, required: [] })).toEqual({
      type: 'object',
      properties: {},
      required: [],
    })
  })

  it('coerces tuple-style items arrays to a single item schema', () => {
    const out = sanitizeGeminiSchema({
      type: 'object',
      properties: { pair: { type: 'array', items: [{ type: 'string' }, { type: 'number' }] } },
    }) as { properties: { pair: { items: unknown } } }
    expect(out.properties.pair.items).toEqual({ type: 'string' })
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

  it('sends Gemini-sanitized tool schemas (type unions collapsed, no additionalProperties)', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okResponse(sseStream([])))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider(
      'gemini',
      { apiKey: 'k', model: 'gemini-2.5-flash' },
      'sys',
      [],
      [
        {
          name: 'apply_patch',
          description: 'patch',
          inputSchema: {
            type: 'object',
            properties: {
              patches: {
                type: 'array',
                items: {
                  type: 'object',
                  properties: { value: { type: ['string', 'null'] }, kind: { type: 'string' } },
                  required: ['kind'],
                },
              },
              extra: { type: 'object', additionalProperties: { type: 'string' } },
            },
            required: ['patches'],
          },
        },
      ],
      100,
      cb,
    )
    const sent = fetchMock.mock.calls[0]![1] as { body: string }
    const body = JSON.parse(sent.body) as {
      tools: Array<{ functionDeclarations: Array<{ parameters: Record<string, unknown> }> }>
    }
    const params = body.tools[0]!.functionDeclarations[0]!.parameters
    const patches = params.properties as Record<string, unknown>
    expect(patches.patches).toMatchObject({
      type: 'array',
      items: {
        type: 'object',
        properties: {
          value: { type: 'string', nullable: true },
          kind: { type: 'string' },
        },
        required: ['kind'],
      },
    })
    expect(patches.extra).not.toHaveProperty('additionalProperties')
    expect(JSON.stringify(body)).not.toContain('additionalProperties')
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

  it('routes openrouter to its fixed base URL', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okResponse(sseStream(['data: [DONE]'])))
    vi.stubGlobal('fetch', fetchMock)
    const { cb } = collector()
    await streamForProvider(
      'openrouter',
      { apiKey: 'sk-or-v1-x', model: 'anthropic/claude-sonnet-4.5' },
      'sys',
      [],
      [],
      100,
      cb,
    )
    expect(fetchMock).toHaveBeenCalledWith(
      'https://openrouter.ai/api/v1/chat/completions',
      expect.objectContaining({
        headers: expect.objectContaining({ Authorization: 'Bearer sk-or-v1-x' }),
      }),
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
})

it('rejects an unknown provider id', async () => {
  const { cb } = collector()
  await expect(
    streamForProvider('unknown' as never, { apiKey: 'k', model: 'm' }, 'sys', [], [], 100, cb),
  ).rejects.toThrow(/Unknown provider/)
})
