import { afterEach, describe, expect, it, vi } from 'vitest'
import type { AgentMessage } from '@genoffice/agent-core'
import { chatForProvider } from '../src/chat'
import { generateImageForProvider } from '../src/image-generation'
import { discoverModels } from '../src/model-discovery'
import { streamForProvider } from '../src/stream'
import { jsonResponse, okResponse, sseStream } from './test-utils'

afterEach(() => {
  vi.unstubAllGlobals()
})

function requestBody(call: unknown[]): Record<string, unknown> {
  const init = call[1] as RequestInit | undefined
  return JSON.parse(String(init?.body)) as Record<string, unknown>
}

function requestHeaders(call: unknown[]): Headers {
  const init = call[1] as RequestInit | undefined
  return new Headers(init?.headers)
}

describe('provider HTTP integration fixtures', () => {
  it('sends OpenAI-compatible chat requests to an overridden endpoint', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ choices: [{ message: { content: 'fixture answer' } }] }))
    vi.stubGlobal('fetch', fetchMock)

    const result = await chatForProvider(
      'openai',
      {
        apiKey: 'sk-fixture',
        model: 'gpt-fixture',
        baseUrl: 'https://llm.example.test/v1/',
      },
      'You are a test assistant.',
      'Answer from the fixture.',
    )

    expect(result).toEqual({ ok: true, content: 'fixture answer' })
    expect(fetchMock).toHaveBeenCalledOnce()
    const call = fetchMock.mock.calls[0] as unknown[]
    expect(call[0]).toBe('https://llm.example.test/v1/chat/completions')
    expect(requestHeaders(call).get('authorization')).toBe('Bearer sk-fixture')
    expect(requestHeaders(call).get('content-type')).toBe('application/json')
    expect(requestBody(call)).toEqual({
      model: 'gpt-fixture',
      messages: [
        { role: 'system', content: 'You are a test assistant.' },
        { role: 'user', content: 'Answer from the fixture.' },
      ],
      temperature: 0.3,
    })
  })

  it('streams OpenAI-compatible text and tool calls through an overridden endpoint', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(
        okResponse(
          sseStream([
            'data: {"choices":[{"delta":{"content":"fixture "}}]}',
            'data: {"choices":[{"delta":{"content":"stream"}}]}',
            'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call-fixture","function":{"name":"lookup"}}]}}]}',
            'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":"{\\"q\\":\\"test\\"}"}}]}}]}',
            'data: {"choices":[{"delta":{},"finish_reason":"tool_calls"}]}',
            'data: [DONE]',
          ]),
        ),
      )
    vi.stubGlobal('fetch', fetchMock)

    const deltas: string[] = []
    const toolCalls: unknown[] = []
    const messages: AgentMessage[] = [{ role: 'user', text: 'lookup test' }]
    await streamForProvider(
      'openrouter',
      {
        apiKey: 'or-fixture',
        model: 'openai/gpt-fixture',
        baseUrl: 'https://router.example.test/api/v1/',
      },
      'Use tools when needed.',
      messages,
      [
        {
          name: 'lookup',
          description: 'Look up a fixture value.',
          inputSchema: { type: 'object', properties: { q: { type: 'string' } } },
        },
      ],
      512,
      {
        signal: new AbortController().signal,
        onDelta: (text) => deltas.push(text),
        onToolCall: (call) => toolCalls.push(call),
      },
    )

    expect(fetchMock).toHaveBeenCalledOnce()
    const call = fetchMock.mock.calls[0] as unknown[]
    expect(call[0]).toBe('https://router.example.test/api/v1/chat/completions')
    expect(requestHeaders(call).get('authorization')).toBe('Bearer or-fixture')
    expect(requestBody(call)).toMatchObject({
      model: 'openai/gpt-fixture',
      max_tokens: 512,
      stream: true,
    })
    expect(deltas.join('')).toBe('fixture stream')
    expect(toolCalls).toEqual([{ id: 'call-fixture', name: 'lookup', input: { q: 'test' } }])
  })

  it('streams Gemini text and function calls from an overridden endpoint', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(
        okResponse(
          sseStream([
            'data: {"candidates":[{"content":{"parts":[{"text":"gemini fixture"}]}}]}',
            'data: {"candidates":[{"content":{"parts":[{"functionCall":{"name":"lookup","args":{"q":"test"}}}]}}]}',
          ]),
        ),
      )
    vi.stubGlobal('fetch', fetchMock)

    const deltas: string[] = []
    const toolCalls: unknown[] = []
    await streamForProvider(
      'gemini',
      {
        apiKey: 'gemini-fixture',
        model: 'gemini-fixture',
        baseUrl: 'https://gemini.example.test/v1beta/',
      },
      'Keep the answer short.',
      [{ role: 'user', text: 'fixture prompt' }],
      [
        {
          name: 'lookup',
          description: 'Look up a fixture value.',
          inputSchema: { type: 'object', properties: { q: { type: 'string' } } },
        },
      ],
      256,
      {
        signal: new AbortController().signal,
        onDelta: (text) => deltas.push(text),
        onToolCall: (call) => toolCalls.push(call),
      },
    )

    const call = fetchMock.mock.calls[0] as unknown[]
    expect(call[0]).toBe(
      'https://gemini.example.test/v1beta/models/gemini-fixture:streamGenerateContent?alt=sse',
    )
    expect(requestHeaders(call).get('x-goog-api-key')).toBe('gemini-fixture')
    expect(requestBody(call)).toMatchObject({
      systemInstruction: { parts: [{ text: 'Keep the answer short.' }] },
      contents: [{ role: 'user', parts: [{ text: 'fixture prompt' }] }],
      generationConfig: { maxOutputTokens: 256 },
    })
    expect(deltas).toEqual(['gemini fixture'])
    expect(toolCalls).toEqual([expect.objectContaining({ name: 'lookup', input: { q: 'test' } })])
  })

  it('discovers and normalizes Anthropic models from an overridden endpoint', async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      jsonResponse({
        data: [
          {
            id: 'claude-fixture',
            display_name: 'Claude Fixture',
            created_at: '2025-01-02T03:04:05Z',
          },
        ],
      }),
    )
    vi.stubGlobal('fetch', fetchMock)

    const catalog = await discoverModels(
      'anthropic',
      {
        apiKey: 'anthropic-fixture',
        model: 'claude-fixture',
        baseUrl: 'https://anthropic.example.test/',
      },
      { timeoutMs: 500 },
    )

    expect(fetchMock).toHaveBeenCalledOnce()
    const call = fetchMock.mock.calls[0] as unknown[]
    expect(call[0]).toBe('https://anthropic.example.test/v1/models?limit=100')
    expect(requestHeaders(call)).toMatchObject(
      new Headers({
        accept: 'application/json',
        'x-api-key': 'anthropic-fixture',
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true',
      }),
    )
    expect(catalog).toMatchObject({
      provider: 'anthropic',
      source: 'remote',
      baseUrl: 'https://anthropic.example.test',
      models: [
        { id: 'claude-fixture', displayName: 'Claude Fixture', capabilities: { chat: true } },
      ],
    })
  })

  it('uses the OpenRouter image endpoint and preserves URL image responses', async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      jsonResponse({
        data: [
          {
            url: 'https://cdn.example.test/fixture.png',
            revised_prompt: 'A revised fixture prompt',
          },
        ],
      }),
    )
    vi.stubGlobal('fetch', fetchMock)

    const images = await generateImageForProvider(
      'openrouter',
      {
        apiKey: 'or-image-fixture',
        model: 'google/gemini-image-fixture',
        baseUrl: 'https://router.example.test/api/v1/',
      },
      { prompt: 'a fixture landscape', count: 1, size: '1024x1024' },
    )

    const call = fetchMock.mock.calls[0] as unknown[]
    expect(call[0]).toBe('https://router.example.test/api/v1/images')
    expect(requestHeaders(call).get('authorization')).toBe('Bearer or-image-fixture')
    expect(requestBody(call)).toEqual({
      model: 'google/gemini-image-fixture',
      prompt: 'a fixture landscape',
      n: 1,
      size: '1024x1024',
    })
    expect(images).toEqual([
      {
        provider: 'openrouter',
        model: 'google/gemini-image-fixture',
        mimeType: 'image/png',
        url: 'https://cdn.example.test/fixture.png',
        revisedPrompt: 'A revised fixture prompt',
      },
    ])
  })

  it('sends Gemini image requests with header auth to an overridden endpoint', async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      jsonResponse({
        candidates: [
          {
            content: {
              parts: [{ inlineData: { mimeType: 'image/jpeg', data: 'Zml4dXJl' } }],
            },
          },
        ],
      }),
    )
    vi.stubGlobal('fetch', fetchMock)

    const images = await generateImageForProvider(
      'gemini',
      {
        apiKey: 'gemini-image-fixture',
        model: 'gemini-image-fixture',
        baseUrl: 'https://gemini.example.test/v1beta/',
      },
      { prompt: 'a fixture icon', aspectRatio: '1:1', imageSize: '2K' },
    )

    const call = fetchMock.mock.calls[0] as unknown[]
    expect(call[0]).toBe(
      'https://gemini.example.test/v1beta/models/gemini-image-fixture:generateContent',
    )
    expect(requestHeaders(call).get('x-goog-api-key')).toBe('gemini-image-fixture')
    expect(requestBody(call)).toMatchObject({
      contents: [{ role: 'user', parts: [{ text: 'a fixture icon' }] }],
      generationConfig: {
        responseModalities: ['TEXT', 'IMAGE'],
        imageConfig: { aspectRatio: '1:1', imageSize: '2K' },
      },
    })
    expect(images).toEqual([
      {
        provider: 'gemini',
        model: 'gemini-image-fixture',
        mimeType: 'image/jpeg',
        base64: 'Zml4dXJl',
      },
    ])
  })
})
