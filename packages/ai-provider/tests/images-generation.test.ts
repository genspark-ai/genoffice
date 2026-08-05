import { afterEach, describe, expect, it, vi } from 'vitest'
import { generateImageForProvider } from '../src/image-generation'
import { jsonResponse } from './test-utils'

afterEach(() => vi.unstubAllGlobals())

describe('generateImageForProvider', () => {
  it('uses OpenAI image generations for OpenAI-compatible providers', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(
        jsonResponse({ data: [{ b64_json: 'aGVsbG8=', revised_prompt: 'revised' }] }),
      )
    vi.stubGlobal('fetch', fetchMock)
    const images = await generateImageForProvider(
      'openai',
      { apiKey: 'sk-test', model: 'gpt-image-1' },
      { prompt: 'a red fox', size: '1024x1024' },
    )
    expect(fetchMock.mock.calls[0]?.[0]).toBe('https://api.openai.com/v1/images/generations')
    const init = fetchMock.mock.calls[0]?.[1] as RequestInit
    expect(JSON.parse(String(init.body))).toMatchObject({
      model: 'gpt-image-1',
      response_format: 'b64_json',
    })
    expect(images[0]).toMatchObject({
      provider: 'openai',
      mimeType: 'image/png',
      base64: 'aGVsbG8=',
      revisedPrompt: 'revised',
    })
  })

  it('uses OpenRouter’s dedicated /api/v1/images contract', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(
        jsonResponse({ data: [{ b64_json: 'aGVsbG8=', media_type: 'image/webp' }] }),
      )
    vi.stubGlobal('fetch', fetchMock)
    await generateImageForProvider(
      'openrouter',
      { apiKey: 'sk-or-test', model: 'google/gemini-3.1-flash-image' },
      { prompt: 'a blue bird' },
    )
    expect(fetchMock.mock.calls[0]?.[0]).toBe('https://openrouter.ai/api/v1/images')
    const init = fetchMock.mock.calls[0]?.[1] as RequestInit
    expect(JSON.parse(String(init.body))).toEqual({
      model: 'google/gemini-3.1-flash-image',
      prompt: 'a blue bird',
    })
  })

  it('parses Gemini inline image parts', async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      jsonResponse({
        candidates: [
          {
            content: {
              parts: [{ inlineData: { mimeType: 'image/png', data: 'aGVsbG8=' } }],
            },
          },
        ],
      }),
    )
    vi.stubGlobal('fetch', fetchMock)
    const images = await generateImageForProvider(
      'gemini',
      { apiKey: 'AIza-test', model: 'gemini-3.1-flash-image' },
      { prompt: 'a green tree', aspectRatio: '1:1', imageSize: '1K' },
    )
    expect(images[0]).toMatchObject({
      provider: 'gemini',
      mimeType: 'image/png',
      base64: 'aGVsbG8=',
    })
    const init = fetchMock.mock.calls[0]?.[1] as RequestInit
    expect(JSON.parse(String(init.body)).generationConfig).toMatchObject({
      responseModalities: ['TEXT', 'IMAGE'],
    })
  })

  it('uses Runware imageInference tasks and returns hosted images', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ data: [{ imageURL: 'https://cdn.runware.ai/test.png' }] }))
    vi.stubGlobal('fetch', fetchMock)

    const images = await generateImageForProvider(
      'runware',
      { apiKey: 'runware-key', model: 'runware:100@1' },
      { prompt: 'editorial watercolor', aspectRatio: '16:9' },
    )

    expect(fetchMock.mock.calls[0]?.[0]).toBe('https://api.runware.ai/v1')
    const init = fetchMock.mock.calls[0]?.[1] as RequestInit
    const body = JSON.parse(String(init.body)) as Array<Record<string, unknown>>
    expect(body[0]).toMatchObject({
      taskType: 'imageInference',
      model: 'runware:100@1',
      positivePrompt: 'editorial watercolor',
      width: 1344,
      height: 768,
      numberResults: 1,
    })
    expect(images[0]?.url).toBe('https://cdn.runware.ai/test.png')
  })

  it('submits a Replicate prediction and parses completed output', async () => {
    const fetchMock = vi.fn().mockResolvedValue(
      jsonResponse({
        status: 'succeeded',
        output: ['https://replicate.delivery/test.webp'],
      }),
    )
    vi.stubGlobal('fetch', fetchMock)

    const images = await generateImageForProvider(
      'replicate',
      { apiKey: 'r8_test', model: 'black-forest-labs/flux-1.1-pro' },
      { prompt: 'a glass office tower', count: 2 },
    )

    expect(fetchMock.mock.calls[0]?.[0]).toBe('https://api.replicate.com/v1/predictions')
    const init = fetchMock.mock.calls[0]?.[1] as RequestInit
    expect(init.headers).toMatchObject({ Authorization: 'Bearer r8_test', Prefer: 'wait=60' })
    expect(JSON.parse(String(init.body))).toMatchObject({
      version: 'black-forest-labs/flux-1.1-pro',
      input: { prompt: 'a glass office tower', num_outputs: 2 },
    })
    expect(images[0]?.url).toBe('https://replicate.delivery/test.webp')
  })

  it('uses fal queue URLs and Key authorization', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(
        jsonResponse({
          status_url: 'https://queue.fal.run/status/1',
          response_url: 'https://queue.fal.run/result/1',
        }),
      )
      .mockResolvedValueOnce(jsonResponse({ status: 'COMPLETED' }))
      .mockResolvedValueOnce(
        jsonResponse({
          images: [{ url: 'https://fal.media/test.png', content_type: 'image/png' }],
        }),
      )
    vi.stubGlobal('fetch', fetchMock)

    const images = await generateImageForProvider(
      'fal',
      { apiKey: 'fal-test', model: 'fal-ai/flux-pro/v1.1-ultra' },
      { prompt: 'technical line art' },
    )

    expect(fetchMock.mock.calls[0]?.[0]).toBe('https://queue.fal.run/fal-ai/flux-pro/v1.1-ultra')
    expect((fetchMock.mock.calls[0]?.[1] as RequestInit).headers).toMatchObject({
      Authorization: 'Key fal-test',
    })
    expect(images[0]?.url).toBe('https://fal.media/test.png')
  })

  it('sends Stability requests as multipart and returns base64 image data', async () => {
    const fetchMock = vi.fn().mockResolvedValue(jsonResponse({ image: 'c3RhYmlsaXR5' }))
    vi.stubGlobal('fetch', fetchMock)

    const images = await generateImageForProvider(
      'stability',
      { apiKey: 'stability-test', model: 'stable-image-core' },
      { prompt: 'minimal geometric poster', aspectRatio: '3:4' },
    )

    expect(fetchMock.mock.calls[0]?.[0]).toBe(
      'https://api.stability.ai/v2beta/stable-image/generate/core',
    )
    const init = fetchMock.mock.calls[0]?.[1] as RequestInit
    expect(init.headers).toMatchObject({
      Authorization: 'Bearer stability-test',
      Accept: 'application/json',
    })
    expect(init.body).toBeInstanceOf(FormData)
    expect((init.body as FormData).get('prompt')).toBe('minimal geometric poster')
    expect(images[0]).toMatchObject({ provider: 'stability', base64: 'c3RhYmlsaXR5' })
  })

  it('rejects unsupported providers and empty prompts', async () => {
    await expect(
      generateImageForProvider('ollama', { apiKey: '', model: 'llama3.2' }, { prompt: 'test' }),
    ).rejects.toThrow('does not advertise image generation')
    await expect(
      generateImageForProvider(
        'openai',
        { apiKey: 'key', model: 'gpt-image-1' },
        { prompt: '   ' },
      ),
    ).rejects.toThrow('Image prompt is required')
  })
})
