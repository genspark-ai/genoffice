import { afterEach, describe, expect, it, vi } from 'vitest'
import { chatForProvider } from '../src/chat'
import { listCustomModels } from '../src/custom-models'
import { AI_PROVIDERS, activeProvider, defaultAiSettings } from '../src/providers'
import { AI_PROVIDER_ADAPTERS } from '../src/registry'
import { jsonResponse } from './test-utils'

afterEach(() => vi.unstubAllGlobals())

const config = (model: string, baseUrl?: string) => ({ apiKey: 'test-key', model, baseUrl })

describe('AI/ML API provider', () => {
  it('resolves to the OpenAI-compatible endpoint on api.aimlapi.com', () => {
    expect(AI_PROVIDER_ADAPTERS.aimlapi.resolveEndpoint(config('openai/gpt-4.1-mini'))).toEqual({
      protocol: 'openai-compatible',
      baseUrl: 'https://api.aimlapi.com/v1',
    })
    expect(AI_PROVIDER_ADAPTERS.aimlapi.capabilities).toEqual({ auth: 'api-key', vision: true })
  })

  it('ignores a stored base URL so the key never leaves api.aimlapi.com', () => {
    for (const baseUrl of [
      'https://evil.example/v1',
      'https://api.aimlapi.com.evil.example/v1',
      'http://api.aimlapi.com/v1',
    ]) {
      expect(
        AI_PROVIDER_ADAPTERS.aimlapi.resolveEndpoint(config('openai/gpt-4.1-mini', baseUrl))
          .baseUrl,
      ).toBe('https://api.aimlapi.com/v1')
    }
  })

  it('sends the chat request with a Bearer key to api.aimlapi.com only', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ choices: [{ message: { content: 'ok' } }] }))
    vi.stubGlobal('fetch', fetchMock)
    const settings = defaultAiSettings()
    settings.provider = 'aimlapi'
    settings.providers.aimlapi = {
      apiKey: 'test-key',
      model: 'openai/gpt-4.1-mini',
      baseUrl: 'https://evil.example/v1',
    }
    expect(activeProvider(settings)).toBe('aimlapi')
    const result = await chatForProvider('aimlapi', settings.providers.aimlapi, 'system', 'hello')
    expect(result).toEqual({ ok: true, content: 'ok' })
    expect(fetchMock).toHaveBeenCalledTimes(1)
    const [url, init] = fetchMock.mock.calls[0]
    expect(url).toBe('https://api.aimlapi.com/v1/chat/completions')
    expect(new URL(url).hostname).toBe('api.aimlapi.com')
    const headers = new Headers(init.headers)
    expect(headers.get('authorization')).toBe('Bearer test-key')
    expect(JSON.parse(init.body).model).toBe('openai/gpt-4.1-mini')
  })

  it('leaves users of other providers on their own provider', () => {
    const settings = defaultAiSettings()
    expect(settings.provider).toBe('genspark')
    expect(activeProvider(settings)).toBe('genspark')
    settings.provider = 'openrouter'
    settings.providers.openrouter.apiKey = 'sk-or-1'
    expect(activeProvider(settings)).toBe('openrouter')
    expect(AI_PROVIDERS.map((p) => p.id)).toContain('aimlapi')
  })
})

describe('AI/ML API live model list', () => {
  const BODY = JSON.stringify({
    data: [
      { id: 'openai/gpt-4.1-mini', type: 'openai/chat-completions' },
      { id: 'anthropic/claude-sonnet-5', type: 'anthropic/messages' },
      { id: 'flux/schnell', type: 'openai/image-generations' },
      { id: 'anthropic/claude-sonnet-5', type: 'openai/chat-completions' },
      { id: 'text-embedding-3-small', type: 'openai/embeddings' },
    ],
  })

  it('keeps only the chat-completions rows, and only from api.aimlapi.com', async () => {
    const fetchMock = vi.fn().mockResolvedValue(jsonResponse(JSON.parse(BODY)))
    vi.stubGlobal('fetch', fetchMock)
    const result = await listCustomModels('https://api.aimlapi.com/v1', 'test-key')
    expect(fetchMock.mock.calls[0][0]).toBe('https://api.aimlapi.com/v1/models')
    expect(result.models).toEqual(['openai/gpt-4.1-mini', 'anthropic/claude-sonnet-5'])
    // a look-alike host is another server: its listing passes through untouched
    fetchMock.mockResolvedValue(jsonResponse(JSON.parse(BODY)))
    const other = await listCustomModels('https://api.aimlapi.com.example/v1', 'test-key')
    expect(other.models).toEqual([
      'openai/gpt-4.1-mini',
      'anthropic/claude-sonnet-5',
      'flux/schnell',
      'text-embedding-3-small',
    ])
  })
})
