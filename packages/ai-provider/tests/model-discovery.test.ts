import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  discoverFalImageModels,
  discoverModels,
  discoverOpenRouterImageModels,
  discoverRunwareImageModels,
  fallbackModelCatalog,
} from '../src/model-discovery'
import { jsonResponse } from './test-utils'

afterEach(() => vi.unstubAllGlobals())

const config = { apiKey: 'test-key', model: 'test-model' }

describe('discoverModels', () => {
  it('reports a user cancellation instead of a generic abort error', async () => {
    const controller = new AbortController()
    vi.stubGlobal(
      'fetch',
      vi.fn((_url: string, init?: RequestInit) => {
        return new Promise<Response>((_resolve, reject) => {
          init?.signal?.addEventListener('abort', () => {
            reject(new DOMException('This operation was aborted', 'AbortError'))
          })
        })
      }),
    )

    const discovery = discoverModels('openai', config, {
      signal: controller.signal,
      timeoutMs: 1_000,
    })
    controller.abort()
    await expect(discovery).rejects.toThrow('Model discovery canceled.')
  })

  it('normalizes OpenAI-compatible model capabilities conservatively', async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      jsonResponse({
        data: [
          {
            id: 'vision-tools-model',
            name: 'Vision + tools',
            created: 1700000000,
            context_length: 131072,
            architecture: { input_modalities: ['text', 'image'], output_modalities: ['text'] },
            supported_parameters: ['tools', 'tool_choice'],
          },
          {
            id: 'plain-model',
            architecture: { input_modalities: ['text'], output_modalities: ['text'] },
          },
        ],
      }),
    )
    vi.stubGlobal('fetch', fetchMock)

    const catalog = await discoverModels('openrouter', config)
    expect(fetchMock).toHaveBeenCalledWith(
      'https://openrouter.ai/api/v1/models',
      expect.objectContaining({
        headers: expect.objectContaining({ Authorization: 'Bearer test-key' }),
      }),
    )
    expect(catalog.source).toBe('remote')
    expect(catalog.models[0]?.capabilities).toEqual({ chat: true, tools: true, vision: true })
    expect(catalog.models[1]?.capabilities).toEqual({ chat: true })
  })

  it('normalizes Gemini names and explicit image output metadata', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        jsonResponse({
          models: [
            {
              name: 'models/gemini-image',
              displayName: 'Gemini Image',
              supportedGenerationMethods: ['generateContent'],
              inputModalities: ['text'],
              outputModalities: ['text', 'image'],
            },
          ],
        }),
      ),
    )
    const catalog = await discoverModels('gemini', config)
    expect(catalog.models[0]).toMatchObject({
      id: 'gemini-image',
      capabilities: { chat: true, imageGeneration: true },
    })
  })

  it('follows every Gemini page and requests the maximum page size', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(
        jsonResponse({
          models: [
            { name: 'models/gemini-chat-a', supportedGenerationMethods: ['generateContent'] },
          ],
          nextPageToken: 'page-two',
        }),
      )
      .mockResolvedValueOnce(
        jsonResponse({
          models: [
            { name: 'models/gemini-chat-b', supportedGenerationMethods: ['generateContent'] },
          ],
        }),
      )
    vi.stubGlobal('fetch', fetchMock)

    const catalog = await discoverModels('gemini', config)
    expect(catalog.models.map((model) => model.id)).toEqual(['gemini-chat-a', 'gemini-chat-b'])
    expect(fetchMock.mock.calls[0]?.[0]).toContain('pageSize=1000')
    expect(fetchMock.mock.calls[1]?.[0]).toContain('pageToken=page-two')
  })

  it('follows Anthropic cursor pages', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(
        jsonResponse({
          data: [{ id: 'claude-first' }],
          has_more: true,
          last_id: 'claude-first',
        }),
      )
      .mockResolvedValueOnce(jsonResponse({ data: [{ id: 'claude-second' }], has_more: false }))
    vi.stubGlobal('fetch', fetchMock)

    const catalog = await discoverModels('anthropic', config)
    expect(catalog.models.map((model) => model.id)).toEqual(['claude-first', 'claude-second'])
    expect(fetchMock.mock.calls[0]?.[0]).toContain('limit=100')
    expect(fetchMock.mock.calls[1]?.[0]).toContain('after_id=claude-first')
  })

  it('uses the dedicated OpenRouter image model catalog', async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      jsonResponse({
        data: [
          {
            id: 'google/gemini-image',
            name: 'Gemini Image',
            architecture: { input_modalities: ['text'], output_modalities: ['image'] },
          },
        ],
      }),
    )
    vi.stubGlobal('fetch', fetchMock)
    const catalog = await discoverOpenRouterImageModels(config)
    expect(fetchMock.mock.calls[0]?.[0]).toBe('https://openrouter.ai/api/v1/images/models')
    expect(catalog.models[0]?.capabilities.imageGeneration).toBe(true)
  })

  it('normalizes Replicate owner and model names from paginated results', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(
        jsonResponse({
          next: 'https://api.replicate.com/v1/models?cursor=next',
          results: [
            {
              owner: 'black-forest-labs',
              name: 'flux-1.1-pro',
              description: 'FLUX 1.1 Pro',
            },
          ],
        }),
      )
      .mockResolvedValueOnce(
        jsonResponse({ results: [{ owner: 'google', name: 'imagen-4-fast' }] }),
      )
    vi.stubGlobal('fetch', fetchMock)
    const catalog = await discoverModels('replicate', config)
    expect(catalog.models[0]).toMatchObject({
      id: 'black-forest-labs/flux-1.1-pro',
      displayName: 'FLUX 1.1 Pro',
      capabilities: { imageGeneration: true },
    })
    expect(catalog.models[1]?.id).toBe('google/imagen-4-fast')
    expect(fetchMock).toHaveBeenCalledTimes(2)
  })

  it('infers image models when OpenAI-compatible catalogs omit modality metadata', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        jsonResponse({
          data: [{ id: 'gpt-4.1' }, { id: 'gpt-image-2' }, { id: 'text-embedding-3-large' }],
        }),
      ),
    )
    const catalog = await discoverModels('openai', config)
    expect(catalog.models.find((model) => model.id === 'gpt-image-2')?.capabilities).toEqual({
      imageGeneration: true,
    })
    expect(catalog.models.find((model) => model.id === 'gpt-4.1')?.capabilities.chat).toBe(true)
    expect(
      catalog.models.find((model) => model.id === 'text-embedding-3-large')?.capabilities.chat,
    ).toBeUndefined()
  })

  it('follows fal model cursors and keeps text-to-image endpoints', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(
        jsonResponse({
          models: [
            {
              endpoint_id: 'fal-ai/flux/dev',
              metadata: { display_name: 'FLUX dev', category: 'text-to-image' },
            },
          ],
          has_more: true,
          next_cursor: 'cursor-two',
        }),
      )
      .mockResolvedValueOnce(
        jsonResponse({
          models: [
            {
              endpoint_id: 'fal-ai/nano-banana-2',
              metadata: { category: 'image-generation' },
            },
          ],
          has_more: false,
          next_cursor: null,
        }),
      )
    vi.stubGlobal('fetch', fetchMock)

    const catalog = await discoverFalImageModels(config)
    expect(catalog.models.map((model) => model.id)).toEqual([
      'fal-ai/flux/dev',
      'fal-ai/nano-banana-2',
    ])
    expect(fetchMock.mock.calls[1]?.[0]).toContain('cursor=cursor-two')
    expect((fetchMock.mock.calls[0]?.[1] as RequestInit).headers).toMatchObject({
      Authorization: 'Key test-key',
    })
  })

  it('paginates Runware merged image models without an unstable visibility filter', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(
        jsonResponse({
          data: [
            {
              results: [{ air: 'runware:100@1', name: 'FLUX Schnell' }],
              totalResults: 2,
            },
          ],
        }),
      )
      .mockResolvedValueOnce(
        jsonResponse({
          data: [
            {
              results: [{ air: 'google:nano-banana@2', name: 'Nano Banana 2' }],
              totalResults: 2,
            },
          ],
        }),
      )
    vi.stubGlobal('fetch', fetchMock)

    const catalog = await discoverRunwareImageModels(config)
    expect(catalog.models.map((model) => model.id)).toEqual([
      'runware:100@1',
      'google:nano-banana@2',
    ])
    const firstRequest = JSON.parse(
      String((fetchMock.mock.calls[0]?.[1] as RequestInit).body),
    ) as Array<Record<string, unknown>>
    const secondRequest = JSON.parse(
      String((fetchMock.mock.calls[1]?.[1] as RequestInit).body),
    ) as Array<Record<string, unknown>>
    expect(firstRequest[0]).toMatchObject({
      taskType: 'modelSearch',
      source: 'merged',
      capabilities: ['text-to-image'],
      offset: 0,
      limit: 100,
    })
    expect(firstRequest[0]).not.toHaveProperty('visibility')
    expect(firstRequest[0]?.source).not.toBe('featured')
    expect(secondRequest[0]?.offset).toBe(100)
  })

  it('limits Runware model probes to one searched page', async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      jsonResponse({
        data: [
          {
            results: [{ air: 'xai:grok-imagine@image-quality', name: 'Grok Imagine' }],
            totalResults: 500,
          },
        ],
      }),
    )
    vi.stubGlobal('fetch', fetchMock)

    const catalog = await discoverRunwareImageModels(config, {
      search: 'xai:grok-imagine@image-quality',
      maxPages: 1,
    })
    expect(fetchMock).toHaveBeenCalledTimes(1)
    const request = JSON.parse(String((fetchMock.mock.calls[0]?.[1] as RequestInit).body)) as Array<
      Record<string, unknown>
    >
    expect(request[0]).toMatchObject({
      search: 'xai:grok-imagine@image-quality',
      offset: 0,
      limit: 100,
    })
    expect(request[0]).not.toHaveProperty('source')
    expect(request[0]).not.toHaveProperty('capabilities')
    expect(catalog.models.map((model) => model.id)).toEqual(['xai:grok-imagine@image-quality'])
  })

  it('returns static fallback models for providers without discovery', () => {
    const fallback = fallbackModelCatalog('genspark')
    expect(fallback.source).toBe('fallback')
    expect(fallback.models.every((model) => model.capabilities.chat)).toBe(true)
  })
})
