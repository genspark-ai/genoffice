import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  GEMINI_FREE_MODELS,
  listModelsForProvider,
  OPENROUTER_FREE_ROUTER,
} from '../src/models'
import { jsonResponse } from './test-utils'

afterEach(() => {
  vi.unstubAllGlobals()
})

const cfg = (apiKey: string) => ({ apiKey, baseUrl: '', model: 'test' })

describe('listModelsForProvider', () => {
  it('lists all OpenRouter models, flagging zero-price and :free ids', async () => {
    const fetchMock = vi.fn(async () =>
      jsonResponse({
        data: [
          { id: 'anthropic/claude-3.7-sonnet', name: 'Claude 3.7', pricing: { prompt: '3', completion: '15' } },
          { id: 'meta-llama/llama-3.3-70b-instruct:free', name: 'Llama 3.3', pricing: { prompt: '0', completion: '0' } },
          { id: 'openai/gpt-4o-mini', pricing: { prompt: '0.15', completion: '0.6' } },
        ],
      }),
    )
    vi.stubGlobal('fetch', fetchMock)
    const models = await listModelsForProvider('openrouter', cfg(''), { freeOnly: false })
    expect((fetchMock.mock.calls[0] as unknown as [string])[0]).toBe(
      'https://openrouter.ai/api/v1/models',
    )
    const free = models.filter((m) => m.free)
    expect(free.map((m) => m.id).sort()).toEqual(['meta-llama/llama-3.3-70b-instruct:free'])
    expect(models.map((m) => m.id)).toContain('openai/gpt-4o-mini')
  })

  it('freeOnly keeps only free models and prepends the free router', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () =>
        jsonResponse({
          data: [
            { id: 'paid/model', pricing: { prompt: '1', completion: '2' } },
            { id: 'free/nano', name: 'Nano', pricing: { prompt: '0', completion: '0' } },
          ],
        }),
      ),
    )
    const models = await listModelsForProvider('openrouter', cfg(''), { freeOnly: true })
    expect(models[0].id).toBe(OPENROUTER_FREE_ROUTER)
    expect(models.filter((m) => m.id !== OPENROUTER_FREE_ROUTER).map((m) => m.id)).toEqual([
      'free/nano',
    ])
  })

  it('throws on a non-200 OpenRouter response', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => new Response('rate limited', { status: 429 })),
    )
    await expect(listModelsForProvider('openrouter', cfg(''))).rejects.toThrow('HTTP 429')
  })

  it('requires a key for Gemini and filters to chat-capable text models', async () => {
    await expect(listModelsForProvider('gemini', cfg(''))).rejects.toThrow('Gemini API key')
    vi.stubGlobal(
      'fetch',
      vi.fn(async () =>
        jsonResponse({
          models: [
            { name: 'models/gemini-2.5-flash', displayName: 'Flash', supportedGenerationMethods: ['generateContent', 'embedContent'] },
            { name: 'models/gemini-3.1-flash-lite', displayName: 'Flash Lite', supportedGenerationMethods: ['generateContent'] },
            { name: 'models/imagen-3.0-generate-002', displayName: 'Imagen', supportedGenerationMethods: ['generateContent'] },
            { name: 'models/gemini-2.5-flash-001', supportedGenerationMethods: ['embedContent'] },
          ],
        }),
      ),
    )
    const models = await listModelsForProvider('gemini', cfg('key'))
    const ids = models.map((m) => m.id)
    expect(ids).toEqual(expect.arrayContaining(['gemini-2.5-flash', 'gemini-3.1-flash-lite']))
    expect(ids).not.toEqual(expect.arrayContaining(['imagen-3.0-generate-002']))
    expect(ids).not.toEqual(expect.arrayContaining(['gemini-2.5-flash-001']))
    expect(models.map((m) => m.id)).toEqual(models.map((m) => m.id).slice().sort())
  })

  it('marks Gemini free-tier routes as free and freeOnly drops the rest', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () =>
        jsonResponse({
          models: [
            { name: 'models/gemini-2.5-flash', displayName: 'Flash', supportedGenerationMethods: ['generateContent'] },
            { name: 'models/gemini-2.5-pro', displayName: 'Pro', supportedGenerationMethods: ['generateContent'] },
          ],
        }),
      ),
    )
    const all = await listModelsForProvider('gemini', cfg('key'), { freeOnly: false })
    expect(all.map((m) => m.id)).toEqual(
      expect.arrayContaining(['gemini-2.5-flash', 'gemini-2.5-pro']),
    )
    expect(all.find((m) => m.id === 'gemini-2.5-flash')?.free).toBe(true)
    const freeOnly = await listModelsForProvider('gemini', cfg('key'), { freeOnly: true })
    expect(freeOnly.map((m) => m.id)).toEqual(['gemini-2.5-flash'])
    expect([...GEMINI_FREE_MODELS]).toContain('gemini-3.5-flash')
  })

  it.each([
    ['anthropic', 'Enter a Claude API key first', 'https://api.anthropic.com/v1/models', [{ id: 'claude-sonnet-4-5', display_name: 'Sonnet' }, { id: 'claude-tools' }]],
    ['openai', 'Enter an OpenAI API key first', 'https://api.openai.com/v1/models', [{ id: 'gpt-4o-mini' }, { id: 'whisper-1' }, { id: 'deepseek-chat' }]],
    ['deepseek', 'Enter a DeepSeek API key first', 'https://api.deepseek.com/models', [{ id: 'deepseek-chat' }, { id: 'deepseek-reasoner' }]],
  ] as const)(
    '%s requires a key, filters non-chat ids, and hits the right endpoint',
    async (
      provider: string,
      errMsg: string,
      url: string,
      data: readonly { id?: string; display_name?: string }[],
    ) => {
      await expect(listModelsForProvider(provider as 'anthropic' | 'openai' | 'deepseek', cfg(''))).rejects.toThrow(errMsg)
      const fetchMock = vi.fn(async () => jsonResponse({ data }))
      vi.stubGlobal('fetch', fetchMock)
      const models = await listModelsForProvider(provider as 'anthropic' | 'openai' | 'deepseek', cfg('key'))
      expect(fetchMock).toHaveBeenCalledWith(url, expect.anything())
      if (provider === 'openai') {
        const ids = models.map((m) => m.id)
        expect(ids).toEqual(['gpt-4o-mini'])
      } else {
        expect(models.length).toBeGreaterThan(0)
      }
    },
  )

  it('genspark returns the static curated list', async () => {
    const models = await listModelsForProvider('genspark', cfg(''))
    expect(models.length).toBeGreaterThan(0)
    expect(models.every((m) => typeof m.id === 'string')).toBe(true)
  })
})