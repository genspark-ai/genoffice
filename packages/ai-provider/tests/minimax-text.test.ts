import { afterEach, describe, expect, it, vi } from 'vitest'
import { chatForProvider } from '../src/chat'
import {
  AI_PROVIDERS,
  activeProvider,
  defaultAiSettings,
  resolveAiSettings,
} from '../src/providers'
import { streamForProvider } from '../src/stream'
import { jsonResponse, okResponse, sseStream } from './test-utils'

afterEach(() => vi.unstubAllGlobals())

describe('MiniMax text choices', () => {
  it('defaults fresh settings to Flash without migrating saved model ids', () => {
    expect(AI_PROVIDERS.find((p) => p.id === 'minimax')?.models).toEqual([
      'MiniMax-M3.1-Flash',
      'MiniMax-M3',
    ])
    expect(defaultAiSettings().providers.minimax.model).toBe('MiniMax-M3.1-Flash')
    for (const model of ['MiniMax-M3', 'MiniMax-M2.7', 'private-model']) {
      const settings = resolveAiSettings(
        { provider: 'minimax', providers: { minimax: { apiKey: 'test-key', model } } },
        defaultAiSettings(),
      )
      expect(settings.providers.minimax.model).toBe(model)
      expect(activeProvider(settings)).toBe('minimax')
    }
    expect(
      AI_PROVIDERS.find((p) => p.id === 'atlascloud')?.models.filter((id) =>
        id.startsWith('minimaxai/'),
      ),
    ).toEqual(['minimaxai/minimax-m3'])
    expect(
      AI_PROVIDERS.find((p) => p.id === 'opencode-zen')?.models.filter((id) =>
        id.startsWith('minimax-'),
      ),
    ).toEqual(['minimax-m3.1-flash-free', 'minimax-m3'])
    expect(
      AI_PROVIDERS.find((p) => p.id === 'opencode-go')?.models.filter((id) =>
        id.startsWith('minimax-'),
      ),
    ).toEqual(['minimax-m3'])
  })

  it.each([undefined, 'https://api.minimaxi.com/v1'])(
    'serializes fresh and saved text requests at the actual fetch boundary (%s)',
    async (baseUrl) => {
      for (const saved of [undefined, 'MiniMax-M3', 'MiniMax-M2.7', 'private-model']) {
        const fetchMock = vi
          .fn()
          .mockResolvedValue(jsonResponse({ choices: [{ message: { content: 'ok' } }] }))
        vi.stubGlobal('fetch', fetchMock)
        const defaults = defaultAiSettings()
        defaults.provider = 'minimax'
        defaults.providers.minimax.apiKey = 'test-key'
        defaults.providers.minimax.baseUrl = baseUrl
        const settings = saved
          ? resolveAiSettings(
              {
                provider: 'minimax',
                providers: { minimax: { ...defaults.providers.minimax, model: saved } },
              },
              defaults,
            )
          : defaults
        const provider = activeProvider(settings)
        expect(
          await chatForProvider(provider, settings.providers[provider], 'system', 'hello'),
        ).toEqual({ ok: true, content: 'ok' })
        const [url, init] = fetchMock.mock.calls[0]
        expect(url).toBe(`${baseUrl ?? 'https://api.minimax.io/v1'}/chat/completions`)
        const body = JSON.parse(init.body)
        expect(body.model).toBe(saved ?? 'MiniMax-M3.1-Flash')
        for (const field of ['thinking', 'reasoning_effort', 'service_tier', 'reasoning_split'])
          expect(body).not.toHaveProperty(field)
      }
    },
  )

  it.each([
    ['opencode-zen', 'minimax-m3.1-flash-free', 'https://opencode.ai/zen/v1/chat/completions'],
    ['opencode-go', 'minimax-m3', 'https://opencode.ai/zen/go/v1/messages'],
    [
      'openrouter',
      'minimax/minimax-m3.1-flash:free',
      'https://openrouter.ai/api/v1/chat/completions',
    ],
    ['atlascloud', 'minimaxai/minimax-m3', 'https://api.atlascloud.ai/v1/chat/completions'],
    ['custom', 'MiniMax-M2.7', 'https://custom.example/v1/chat/completions'],
  ] as const)('keeps %s model ids and transport contracts', async (provider, model, url) => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(
        jsonResponse(
          provider === 'opencode-go'
            ? { content: [{ type: 'text', text: 'ok' }] }
            : { choices: [{ message: { content: 'ok' } }] },
        ),
      )
    vi.stubGlobal('fetch', fetchMock)
    const defaults = defaultAiSettings()
    const settings = resolveAiSettings(
      {
        provider,
        providers: {
          [provider]: {
            apiKey: 'test-key',
            model,
            ...(provider === 'custom' ? { baseUrl: 'https://custom.example/v1' } : {}),
          },
        },
      },
      defaults,
    )
    expect(
      await chatForProvider(
        activeProvider(settings),
        settings.providers[provider],
        'system',
        'hello',
      ),
    ).toEqual({ ok: true, content: 'ok' })
    expect(fetchMock.mock.calls[0][0]).toBe(url)
    expect(JSON.parse(fetchMock.mock.calls[0][1].body).model).toBe(model)
  })

  it.each([undefined, 'https://api.minimaxi.com/v1'])(
    'streams the default model with images and tool history intact (%s)',
    async (baseUrl) => {
      const fetchMock = vi
        .fn()
        .mockResolvedValue(
          okResponse(
            sseStream([
              'data: {"choices":[{"delta":{"content":"ok"},"finish_reason":"stop"}]}',
              'data: [DONE]',
            ]),
          ),
        )
      vi.stubGlobal('fetch', fetchMock)
      const settings = defaultAiSettings()
      settings.provider = 'minimax'
      settings.providers.minimax.apiKey = 'test-key'
      settings.providers.minimax.baseUrl = baseUrl
      await streamForProvider(
        activeProvider(settings),
        settings.providers.minimax,
        'system',
        [
          { role: 'user', text: 'Describe', images: [{ base64: 'AQID', mime: 'image/png' }] },
          {
            role: 'assistant',
            text: '',
            reasoning: 'Inspect the image',
            toolCalls: [{ id: 'call-1', name: 'inspect', input: {} }],
          },
          { role: 'tool', results: [{ id: 'call-1', name: 'inspect', output: 'done' }] },
        ],
        [
          {
            name: 'inspect',
            description: 'Inspect',
            inputSchema: { type: 'object', properties: {} },
          },
        ],
        32768,
        { signal: new AbortController().signal, onDelta: () => {}, onToolCall: () => {} },
      )
      const [url, init] = fetchMock.mock.calls[0]
      expect(url).toBe(`${baseUrl ?? 'https://api.minimax.io/v1'}/chat/completions`)
      const body = JSON.parse(init.body)
      expect(body).toMatchObject({ model: 'MiniMax-M3.1-Flash', stream: true, max_tokens: 32768 })
      expect(body.messages[1].content).toContainEqual({
        type: 'image_url',
        image_url: { url: 'data:image/png;base64,AQID' },
      })
      expect(body.messages[2].reasoning_content).toBe('Inspect the image')
      expect(body.messages[3]).toMatchObject({
        role: 'tool',
        tool_call_id: 'call-1',
        content: 'done',
      })
      for (const field of ['thinking', 'reasoning_effort', 'service_tier', 'reasoning_split'])
        expect(body).not.toHaveProperty(field)
    },
  )
})
