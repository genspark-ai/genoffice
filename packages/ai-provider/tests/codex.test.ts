import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  buildCodexRequest,
  codexHeaders,
  fetchCodexCapabilities,
  streamCodexResponse,
} from '../src/codex'
import { setRescueFetch } from '../src/fetch'
import { streamForProvider } from '../src/stream'
import { AI_CONNECT_TIMEOUT_MS, AiTimeoutError } from '../src/watchdog'
import { okResponse, sseStream } from './test-utils'

const auth = {
  accessToken: 'opaque-access-token',
  accountId: 'account-123',
  expiresAt: 1_800_000_000_000,
}

const incompleteStreams: Array<[string, string[]]> = [
  ['EOF', []],
  ['[DONE]', ['data: [DONE]']],
]

afterEach(() => {
  vi.useRealTimers()
  vi.unstubAllGlobals()
  setRescueFetch(null)
})

describe('buildCodexRequest', () => {
  it('translates instructions, history, images, tools, and tool results into Responses input', () => {
    const body = buildCodexRequest({
      auth,
      instructions: 'Edit only requested document content.',
      model: 'gpt-5.5',
      signal: new AbortController().signal,
      onDelta: () => {},
      onToolCall: () => {},
      tools: [
        {
          name: 'docs_replace',
          description: 'Replace document text.',
          inputSchema: { type: 'object', properties: { text: { type: 'string' } } },
        },
      ],
      messages: [
        {
          role: 'user',
          text: 'Replace title',
          images: [{ mime: 'image/png', base64: 'image-data' }],
        },
        {
          role: 'assistant',
          text: 'I will update it.',
          toolCalls: [{ id: 'call-1', name: 'docs_replace', input: { text: 'New title' } }],
        },
        {
          role: 'tool',
          results: [{ id: 'call-1', name: 'docs_replace', output: 'Updated 1 block.' }],
        },
      ],
    })

    expect(body).toEqual({
      model: 'gpt-5.5',
      instructions: 'Edit only requested document content.',
      input: [
        {
          role: 'user',
          content: [
            { type: 'input_text', text: 'Replace title' },
            { type: 'input_image', image_url: 'data:image/png;base64,image-data' },
          ],
        },
        { role: 'assistant', content: [{ type: 'output_text', text: 'I will update it.' }] },
        {
          type: 'function_call',
          call_id: 'call-1',
          name: 'docs_replace',
          arguments: '{"text":"New title"}',
        },
        { type: 'function_call_output', call_id: 'call-1', output: 'Updated 1 block.' },
      ],
      tools: [
        {
          type: 'function',
          name: 'docs_replace',
          description: 'Replace document text.',
          parameters: { type: 'object', properties: { text: { type: 'string' } } },
          strict: false,
        },
      ],
      store: false,
      stream: true,
    })
  })

  it('sends opaque account credentials only as request headers', () => {
    expect(codexHeaders(auth, 'session-123')).toEqual({
      Accept: 'text/event-stream',
      Authorization: 'Bearer opaque-access-token',
      'Content-Type': 'application/json',
      'OpenAI-Beta': 'responses=experimental',
      'chatgpt-account-id': 'account-123',
      originator: 'codex_cli_rs',
      session_id: 'session-123',
    })
  })

  it.each(['minimal', 'low', 'medium', 'high', 'xhigh', 'max', 'ultra'] as const)(
    'serializes %s reasoning effort',
    (reasoningEffort) => {
      expect(
        buildCodexRequest({
          auth,
          instructions: '',
          model: 'gpt-test',
          reasoningEffort,
          messages: [],
          tools: [],
          signal: new AbortController().signal,
          onDelta: () => {},
          onToolCall: () => {},
        }),
      ).toMatchObject({ reasoning: { effort: reasoningEffort } })
    },
  )

  it('omits none reasoning effort', () => {
    expect(
      buildCodexRequest({
        auth,
        instructions: '',
        model: 'gpt-test',
        reasoningEffort: 'none',
        messages: [],
        tools: [],
        signal: new AbortController().signal,
        onDelta: () => {},
        onToolCall: () => {},
      }),
    ).not.toHaveProperty('reasoning')
  })

  it('serializes a non-default service tier and omits the default tier', () => {
    expect(
      buildCodexRequest({
        auth,
        instructions: '',
        model: 'gpt-test',
        serviceTier: 'priority',
        messages: [],
        tools: [],
        signal: new AbortController().signal,
        onDelta: () => {},
        onToolCall: () => {},
      }),
    ).toMatchObject({ service_tier: 'priority' })
    expect(
      buildCodexRequest({
        auth,
        instructions: '',
        model: 'gpt-test',
        serviceTier: 'default',
        messages: [],
        tools: [],
        signal: new AbortController().signal,
        onDelta: () => {},
        onToolCall: () => {},
      }),
    ).not.toHaveProperty('service_tier')
  })
})

describe('fetchCodexCapabilities', () => {
  it('returns only picker-visible API models and validated reasoning efforts', async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      new Response(
        JSON.stringify({
          models: [
            {
              slug: 'gpt-5.5',
              display_name: 'GPT-5.5',
              visibility: 'list',
              supported_in_api: true,
              supported_reasoning_levels: [
                { effort: 'low' },
                { effort: 'high' },
                { effort: 'bad' },
              ],
              service_tiers: [{ id: 'priority', name: 'Fast', description: '1.5x speed' }],
              default_service_tier: 'default',
            },
            {
              slug: 'hidden',
              visibility: 'hide',
              supported_in_api: true,
              supported_reasoning_levels: [{ effort: 'medium' }],
            },
            {
              slug: 'gpt-5.5',
              visibility: 'list',
              supported_in_api: true,
              supported_reasoning_levels: [{ effort: 'ultra' }],
            },
          ],
        }),
      ),
    )

    await expect(fetchCodexCapabilities(auth, undefined, fetchMock)).resolves.toEqual({
      models: [
        {
          id: 'gpt-5.5',
          name: '5.5',
          reasoningEfforts: ['low', 'high'],
          serviceTiers: [
            { id: 'default', name: 'Standard' },
            { id: 'priority', name: 'Fast', description: '1.5x speed' },
          ],
          defaultServiceTier: 'default',
        },
      ],
    })
    expect(fetchMock.mock.calls[0]?.[0]).toContain('/codex/models?client_version=0.144.1')
  })

  it('rejects malformed and empty catalogs without retaining their body', async () => {
    await expect(
      fetchCodexCapabilities(
        auth,
        undefined,
        vi.fn().mockResolvedValue(new Response('{"models":"bad"}')),
      ),
    ).rejects.toMatchObject({ code: 'capabilities-unavailable' })
    await expect(
      fetchCodexCapabilities(
        auth,
        undefined,
        vi.fn().mockResolvedValue(new Response('{"models":[]}')),
      ),
    ).rejects.toMatchObject({ code: 'capabilities-unavailable' })
  })
})

describe('streamCodexResponse', () => {
  it('emits text and complete parallel tool calls without executing them', async () => {
    vi.stubGlobal(
      'fetch',
      vi
        .fn()
        .mockResolvedValue(
          okResponse(
            sseStream([
              'data: {"type":"response.output_text.delta","delta":"Draft "}',
              'data: {"type":"response.output_item.added","item":{"type":"function_call","call_id":"call-1","name":"docs_replace"}}',
              'data: {"type":"response.output_item.added","item":{"type":"function_call","call_id":"call-2","name":"files_read"}}',
              'data: {"type":"response.function_call_arguments.delta","call_id":"call-1","delta":"{\\"text\\":\\"New"}',
              'data: {"type":"response.function_call_arguments.delta","call_id":"call-2","delta":"{\\"path\\":\\"notes.txt\\"}"}',
              'data: {"type":"response.function_call_arguments.delta","call_id":"call-1","delta":" title\\"}"}',
              'data: {"type":"response.function_call_arguments.done","call_id":"call-2"}',
              'data: {"type":"response.function_call_arguments.done","call_id":"call-1"}',
              'data: {"type":"response.completed"}',
            ]),
          ),
        ),
    )
    const deltas: string[] = []
    const calls: Array<{ id: string; name: string; input: Record<string, unknown> }> = []

    await streamCodexResponse({
      auth,
      instructions: 'Write.',
      model: 'gpt-5.5',
      messages: [{ role: 'user', text: 'Start' }],
      tools: [],
      signal: new AbortController().signal,
      onDelta: (delta) => deltas.push(delta),
      onToolCall: (call) => calls.push(call),
    })

    expect(deltas).toEqual(['Draft '])
    expect(calls).toEqual([
      { id: 'call-2', name: 'files_read', input: { path: 'notes.txt' } },
      { id: 'call-1', name: 'docs_replace', input: { text: 'New title' } },
    ])
  })

  it.each(incompleteStreams)(
    'rejects a stream ending with %s before response.completed',
    async (_ending, lines) => {
      vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(sseStream(lines))))

      await expect(
        streamCodexResponse({
          auth,
          instructions: 'Write.',
          model: 'gpt-5.5',
          messages: [],
          tools: [],
          signal: new AbortController().signal,
          onDelta: () => {},
          onToolCall: () => {},
        }),
      ).rejects.toMatchObject({ code: 'invalid-stream' })
    },
  )

  it('reports received bytes through onActivity', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(okResponse(sseStream(['data: {"type":"response.completed"}']))),
    )
    const onActivity = vi.fn()

    await streamCodexResponse({
      auth,
      instructions: 'Write.',
      model: 'gpt-5.5',
      messages: [],
      tools: [],
      signal: new AbortController().signal,
      onDelta: () => {},
      onToolCall: () => {},
      onActivity,
    })

    expect(onActivity).toHaveBeenCalled()
  })

  it('retries an initial network failure through the rescue fetch', async () => {
    const primary = vi.fn().mockRejectedValue(new Error('network failure'))
    const rescue = vi
      .fn()
      .mockResolvedValue(okResponse(sseStream(['data: {"type":"response.completed"}'])))
    vi.stubGlobal('fetch', primary)
    setRescueFetch(rescue)

    await streamCodexResponse({
      auth,
      instructions: 'Write.',
      model: 'gpt-5.5',
      messages: [],
      tools: [],
      signal: new AbortController().signal,
      onDelta: () => {},
      onToolCall: () => {},
    })

    expect(primary).toHaveBeenCalledOnce()
    expect(rescue).toHaveBeenCalledOnce()
  })

  it('reports watchdog expiry as AiTimeoutError', async () => {
    vi.useFakeTimers()
    vi.stubGlobal(
      'fetch',
      vi.fn(
        (_url: string, init: RequestInit) =>
          new Promise((_, reject) => {
            init.signal?.addEventListener('abort', () => reject(new Error('aborted')), {
              once: true,
            })
          }),
      ),
    )

    const run = streamCodexResponse({
      auth,
      instructions: 'Write.',
      model: 'gpt-5.5',
      messages: [],
      tools: [],
      signal: new AbortController().signal,
      onDelta: () => {},
      onToolCall: () => {},
    })
    const result = expect(run).rejects.toBeInstanceOf(AiTimeoutError)
    await vi.advanceTimersByTimeAsync(AI_CONNECT_TIMEOUT_MS)
    await result
  })

  it('raises a bounded error for malformed tool arguments', async () => {
    vi.stubGlobal(
      'fetch',
      vi
        .fn()
        .mockResolvedValue(
          okResponse(
            sseStream([
              'data: {"type":"response.output_item.added","item":{"type":"function_call","call_id":"call-1","name":"docs_replace"}}',
              'data: {"type":"response.function_call_arguments.delta","call_id":"call-1","delta":"{broken"}',
              'data: {"type":"response.function_call_arguments.done","call_id":"call-1"}',
            ]),
          ),
        ),
    )

    await expect(
      streamCodexResponse({
        auth,
        instructions: 'Write.',
        model: 'gpt-5.5',
        messages: [],
        tools: [],
        signal: new AbortController().signal,
        onDelta: () => {},
        onToolCall: () => {},
      }),
    ).rejects.toMatchObject({ code: 'invalid-tool-call' })
  })

  it.each([
    [401, 'auth-expired'],
    [429, 'rate-limit'],
    [500, 'provider-failure'],
  ] as const)('normalizes Codex HTTP %s into a provider error code', async (status, code) => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response('backend detail', { status })))

    await expect(
      streamCodexResponse({
        auth,
        instructions: 'Write.',
        model: 'gpt-5.5',
        messages: [],
        tools: [],
        signal: new AbortController().signal,
        onDelta: () => {},
        onToolCall: () => {},
      }),
    ).rejects.toMatchObject({ name: 'CodexHttpError', status, code })
  })

  it('classifies a safe 400 diagnostic without retaining provider text', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        new Response(
          JSON.stringify({
            error: { code: 'invalid_model', message: 'The selected model is unavailable.' },
          }),
          { status: 400 },
        ),
      ),
    )

    await expect(
      streamCodexResponse({
        auth,
        instructions: 'Write.',
        model: 'gpt-5.5',
        messages: [],
        tools: [],
        signal: new AbortController().signal,
        onDelta: () => {},
        onToolCall: () => {},
      }),
    ).rejects.toMatchObject({
      name: 'CodexHttpError',
      status: 400,
      code: 'capabilities-unavailable',
      diagnosticCode: 'invalid_model',
    })
  })

  it('does not retain credentials or URLs from a 400 response body', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        new Response(
          JSON.stringify({
            error: {
              code: 'invalid_request',
              message: 'Bad request. token=secret-token https://private.example/path',
            },
          }),
          { status: 400 },
        ),
      ),
    )

    const error = await streamCodexResponse({
      auth,
      instructions: 'Write.',
      model: 'gpt-5.5',
      messages: [],
      tools: [],
      signal: new AbortController().signal,
      onDelta: () => {},
      onToolCall: () => {},
    }).catch((reason: unknown) => reason)

    expect(error).toMatchObject({
      name: 'CodexHttpError',
      status: 400,
      code: 'request-rejected',
      diagnosticCode: 'invalid_request',
    })
    expect(String(error)).not.toContain('secret-token')
    expect(String(error)).not.toContain('private.example')
  })

  it('rejects malformed SSE JSON before emitting a callback', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(okResponse(sseStream(['data: {not-json}']))))

    await expect(
      streamCodexResponse({
        auth,
        instructions: 'Write.',
        model: 'gpt-5.5',
        messages: [],
        tools: [],
        signal: new AbortController().signal,
        onDelta: () => {},
        onToolCall: () => {},
      }),
    ).rejects.toMatchObject({ code: 'invalid-stream' })
  })

  it('passes cancellation signal to fetch and stops before stale callbacks', async () => {
    const controller = new AbortController()
    const fetchMock = vi
      .fn()
      .mockResolvedValue(
        okResponse(sseStream(['data: {"type":"response.output_text.delta","delta":"late"}'])),
      )
    vi.stubGlobal('fetch', fetchMock)
    const deltas: string[] = []
    controller.abort()

    await expect(
      streamCodexResponse({
        auth,
        instructions: 'Write.',
        model: 'gpt-5.5',
        messages: [],
        tools: [],
        signal: controller.signal,
        onDelta: (delta) => deltas.push(delta),
        onToolCall: () => {},
      }),
    ).rejects.toMatchObject({ name: 'AbortError' })
    expect(deltas).toEqual([])
    expect(fetchMock).not.toHaveBeenCalled()
  })

  it('routes Codex through only caller-supplied Docs tools', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(okResponse(sseStream(['data: {"type":"response.completed"}'])))
    vi.stubGlobal('fetch', fetchMock)

    await streamForProvider(
      'openai-codex',
      { apiKey: '', model: 'gpt-5.5' },
      'Edit the document.',
      [{ role: 'user', text: 'Replace title' }],
      [
        {
          name: 'docs_replace',
          description: 'Replace document text.',
          inputSchema: { type: 'object' },
        },
      ],
      100,
      { signal: new AbortController().signal, onDelta: () => {}, onToolCall: () => {} },
      auth,
    )

    const body = JSON.parse(fetchMock.mock.calls[0]![1].body) as { tools: Array<{ name: string }> }
    expect(body.tools).toEqual([
      {
        type: 'function',
        name: 'docs_replace',
        description: 'Replace document text.',
        parameters: { type: 'object' },
        strict: false,
      },
    ])
  })
})
