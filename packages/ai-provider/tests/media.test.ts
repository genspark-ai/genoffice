import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  AI_MEDIA_PROVIDERS,
  activeMediaConfig,
  activeMediaProvider,
  defaultAiMediaSettings,
  imageGenerationAvailable,
  mediaAnalysisAvailable,
  resolveAiMediaSettings,
} from '../src/media'
import {
  analyzeMediaWithProvider,
  generateImageWithProvider,
  sniffImageMime,
  testMediaProvider,
} from '../src/media-protocols'
import { defaultAiSettings, resolveAiSettings } from '../src/providers'
import type { AiSettings } from '../src/types'

afterEach(() => {
  vi.unstubAllGlobals()
})

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0])
const PNG_B64 = Buffer.from(PNG).toString('base64')

function withMedia(media: Partial<AiSettings['media']>): AiSettings {
  const base = defaultAiSettings()
  return { ...base, media: { ...base.media!, ...media } as AiSettings['media'] }
}

function openaiSettings(apiKey = 'sk-test', imageModel = 'gpt-image-2'): AiSettings {
  const media = defaultAiMediaSettings()
  media.imageProvider = 'openai'
  media.analysisProvider = 'openai'
  media.providers.openai = { apiKey, imageModel, analysisModel: 'gpt-5.6-luna' }
  return withMedia(media)
}

describe('media settings', () => {
  it('defaults every provider to its default models and genspark as the active one', () => {
    const media = defaultAiMediaSettings()
    expect(media.imageProvider).toBe('genspark')
    expect(media.analysisProvider).toBe('genspark')
    expect(media.videoAnalysisProvider).toBe('genspark')
    for (const meta of AI_MEDIA_PROVIDERS) {
      expect(media.providers[meta.id].imageModel).toBe(meta.defaultImageModel)
      expect(media.providers[meta.id].apiKey).toBe('')
    }
    expect(media.providers.custom.baseUrl).toBe('')
    expect(media.providers.openai.baseUrl).toBeUndefined()
  })

  it('is carried by defaultAiSettings and healed in from a pre-media settings file', () => {
    const defaults = defaultAiSettings()
    expect(defaults.media?.imageProvider).toBe('genspark')
    const resolved = resolveAiSettings(
      { provider: 'genspark', providers: defaults.providers },
      defaultAiSettings(),
    )
    expect(resolved.media).toEqual(defaultAiMediaSettings())
  })

  it('merges and trims a stored media block over defaults', () => {
    const media = resolveAiMediaSettings({
      provider: 'gemini',
      providers: {
        gemini: { apiKey: ' AIza ', imageModel: ' gemini-3-pro-image ', analysisModel: '' },
      } as never,
    })
    // the pre-catalog single `provider` seeds both capabilities
    expect(media.imageProvider).toBe('gemini')
    expect(media.analysisProvider).toBe('gemini')
    expect(
      resolveAiMediaSettings({ imageProvider: 'doubao', analysisProvider: 'qwen' })
        .analysisProvider,
    ).toBe('qwen')
    // pre-split files used one vendor for all analysis
    expect(
      resolveAiMediaSettings({ imageProvider: 'doubao', analysisProvider: 'qwen' })
        .videoAnalysisProvider,
    ).toBe('qwen')
    expect(
      resolveAiMediaSettings({ analysisProvider: 'openai', videoAnalysisProvider: 'gemini' })
        .videoAnalysisProvider,
    ).toBe('gemini')
    expect(media.providers.gemini).toEqual({
      apiKey: 'AIza',
      imageModel: 'gemini-3-pro-image',
      analysisModel: '',
    })
    expect(media.providers.openai.imageModel).toBe('gpt-image-2')
  })

  it('activates a BYOK media provider per capability, only when usable and capable', () => {
    expect(activeMediaProvider(openaiSettings(), 'image')).toBe('openai')
    expect(activeMediaProvider(openaiSettings(), 'analysis')).toBe('openai')
    expect(activeMediaProvider(openaiSettings(''), 'image')).toBe('genspark')
    const custom = defaultAiMediaSettings()
    custom.imageProvider = 'custom'
    expect(activeMediaProvider(withMedia(custom), 'image')).toBe('genspark')
    custom.providers.custom.baseUrl = 'http://localhost:1234/v1'
    expect(activeMediaProvider(withMedia(custom), 'image')).toBe('custom')
    expect(activeMediaProvider(withMedia(custom), 'analysis')).toBe('genspark')
    expect(activeMediaConfig(withMedia(custom), 'image')?.provider).toBe('custom')
    // MiniMax has no analysis endpoint: picking it for analysis falls back
    const mm = defaultAiMediaSettings()
    mm.analysisProvider = 'minimax'
    mm.providers.minimax.apiKey = 'k'
    expect(activeMediaProvider(withMedia(mm), 'analysis')).toBe('genspark')
    // OpenAI reads images but not video: as the video provider it falls back
    const oa = openaiSettings()
    oa.media!.videoAnalysisProvider = 'openai'
    expect(activeMediaProvider(oa, 'video')).toBe('genspark')
    oa.media!.videoAnalysisProvider = 'gemini'
    oa.media!.providers.gemini.apiKey = 'AIza'
    expect(activeMediaProvider(oa, 'video')).toBe('gemini')
    expect(activeMediaProvider({ media: undefined }, 'image')).toBe('genspark')
    expect(
      activeMediaProvider({ media: { imageProvider: 'nope', providers: {} } as never }, 'image'),
    ).toBe('genspark')
  })

  it('gates the tools on gsk login + toggle without BYOK, and on the BYOK model with it', () => {
    const genspark = defaultAiSettings()
    expect(imageGenerationAvailable(genspark, true)).toBe(true)
    expect(imageGenerationAvailable(genspark, false)).toBe(false)
    expect(imageGenerationAvailable({ ...genspark, gskToolsEnabled: false }, true)).toBe(false)
    expect(mediaAnalysisAvailable({ ...genspark, gskToolsEnabled: false }, true)).toBe(false)

    const byok = openaiSettings()
    expect(imageGenerationAvailable(byok, false)).toBe(true)
    expect(imageGenerationAvailable({ ...byok, gskToolsEnabled: false }, false)).toBe(true)
    expect(mediaAnalysisAvailable(byok, false)).toBe(true)
    // built-in vendors fall back to their default model; a custom endpoint
    // without a model for one capability disables that tool alone
    const noAnalysis = openaiSettings()
    noAnalysis.media!.providers.openai.analysisModel = ''
    expect(mediaAnalysisAvailable(noAnalysis, false)).toBe(true)
    const custom = defaultAiMediaSettings()
    custom.imageProvider = 'custom'
    custom.analysisProvider = 'custom'
    custom.providers.custom = {
      apiKey: '',
      baseUrl: 'http://localhost:1234/v1',
      imageModel: 'flux',
      analysisModel: '',
    }
    expect(mediaAnalysisAvailable(withMedia(custom), false)).toBe(false)
    expect(imageGenerationAvailable(withMedia(custom), false)).toBe(true)
    expect(imageGenerationAvailable(null, true)).toBe(true)
  })
})

describe('sniffImageMime', () => {
  it('reads the magic bytes and falls back to the declared type', () => {
    expect(sniffImageMime(PNG)).toBe('image/png')
    expect(sniffImageMime(new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0]))).toBe('image/jpeg')
    expect(sniffImageMime(new Uint8Array([1, 2, 3]), 'image/webp')).toBe('image/webp')
  })
})

function jsonResponse(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json' },
  })
}

describe('generateImageWithProvider', () => {
  it('posts to the OpenAI Images API with a size derived from the aspect ratio and decodes b64_json', async () => {
    const fetchMock = vi.fn(async () => jsonResponse({ data: [{ b64_json: PNG_B64 }] }))
    vi.stubGlobal('fetch', fetchMock)
    const out = await generateImageWithProvider(
      'openai',
      { apiKey: 'sk-1', imageModel: 'gpt-image-2', analysisModel: '' },
      { prompt: 'a cat', aspectRatio: '16:9' },
    )
    expect(out.mime).toBe('image/png')
    expect(Array.from(out.bytes)).toEqual(Array.from(PNG))
    const [url, init] = fetchMock.mock.calls[0] as unknown as [string, RequestInit]
    expect(url).toBe('https://api.openai.com/v1/images/generations')
    expect((init.headers as Record<string, string>).Authorization).toBe('Bearer sk-1')
    expect(JSON.parse(init.body as string)).toEqual({
      model: 'gpt-image-2',
      prompt: 'a cat',
      n: 1,
      size: '1536x1024',
    })
  })

  it('uses the edits endpoint (multipart) when reference images are given', async () => {
    const fetchMock = vi.fn(async () => jsonResponse({ data: [{ b64_json: PNG_B64 }] }))
    vi.stubGlobal('fetch', fetchMock)
    await generateImageWithProvider(
      'custom',
      { apiKey: '', baseUrl: 'http://localhost:1234/v1/', imageModel: 'flux', analysisModel: '' },
      { prompt: 'remove background', references: [{ bytes: PNG, mime: 'image/png' }] },
    )
    const [url, init] = fetchMock.mock.calls[0] as unknown as [string, RequestInit]
    expect(url).toBe('http://localhost:1234/v1/images/edits')
    expect(init.body).toBeInstanceOf(FormData)
    const form = init.body as FormData
    expect(form.get('model')).toBe('flux')
    expect(form.get('image')).toBeInstanceOf(Blob)
    expect((init.headers as Record<string, string>).Authorization).toBeUndefined()
  })

  it('calls Gemini generateContent with IMAGE modality and reads inlineData back', async () => {
    const fetchMock = vi.fn(async () =>
      jsonResponse({
        candidates: [
          {
            content: {
              parts: [
                { text: 'here you go' },
                { inlineData: { mimeType: 'image/png', data: PNG_B64 } },
              ],
            },
          },
        ],
      }),
    )
    vi.stubGlobal('fetch', fetchMock)
    const out = await generateImageWithProvider(
      'gemini',
      { apiKey: 'AIza', imageModel: 'gemini-3.1-flash-image', analysisModel: '' },
      { prompt: 'a dog', aspectRatio: '9:16' },
    )
    expect(out.mime).toBe('image/png')
    const [url, init] = fetchMock.mock.calls[0] as unknown as [string, RequestInit]
    expect(url).toBe(
      'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-image:generateContent',
    )
    expect((init.headers as Record<string, string>)['x-goog-api-key']).toBe('AIza')
    const body = JSON.parse(init.body as string)
    expect(body.generationConfig).toEqual({
      responseModalities: ['TEXT', 'IMAGE'],
      imageConfig: { aspectRatio: '9:16' },
    })
  })

  it('routes imagen-* ids to :predict', async () => {
    const fetchMock = vi.fn(async () =>
      jsonResponse({ predictions: [{ bytesBase64Encoded: PNG_B64, mimeType: 'image/png' }] }),
    )
    vi.stubGlobal('fetch', fetchMock)
    await generateImageWithProvider(
      'gemini',
      { apiKey: 'AIza', imageModel: 'imagen-4.0-generate-001', analysisModel: '' },
      { prompt: 'a dog' },
    )
    const [url] = fetchMock.mock.calls[0] as unknown as [string]
    expect(url).toMatch(/imagen-4\.0-generate-001:predict$/)
  })

  it('surfaces HTTP errors with the body detail and refuses custom without a base URL', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => jsonResponse({ error: { message: 'invalid api key' } }, 401)),
    )
    await expect(
      generateImageWithProvider(
        'openai',
        { apiKey: 'bad', imageModel: 'gpt-image-2', analysisModel: '' },
        { prompt: 'x' },
      ),
    ).rejects.toThrow(/401.*invalid api key/)
    await expect(
      generateImageWithProvider(
        'custom',
        { apiKey: '', imageModel: 'm', analysisModel: '' },
        { prompt: 'x' },
      ),
    ).rejects.toThrow(/Base URL/)
  })
})

describe('generateImageWithProvider (vendor shapes)', () => {
  const cfg = (imageModel: string, baseUrl?: string) => ({
    apiKey: 'k',
    imageModel,
    analysisModel: '',
    ...(baseUrl ? { baseUrl } : {}),
  })

  it('Ark/Seedream: aspect ratio as size, b64 requested, references inline in the body', async () => {
    const fetchMock = vi.fn(async () => jsonResponse({ data: [{ b64_json: PNG_B64 }] }))
    vi.stubGlobal('fetch', fetchMock)
    await generateImageWithProvider('doubao', cfg('doubao-seedream-5-0-260128'), {
      prompt: 'a fox',
      aspectRatio: '16:9',
      references: [{ bytes: PNG, mime: 'image/png' }],
    })
    const [url, init] = fetchMock.mock.calls[0] as unknown as [string, RequestInit]
    expect(url).toBe('https://ark.cn-beijing.volces.com/api/v3/images/generations')
    const body = JSON.parse(init.body as string)
    expect(body).toMatchObject({ size: '16:9', response_format: 'b64_json', watermark: false })
    expect(body.image).toEqual([`data:image/png;base64,${PNG_B64}`])
  })

  it('Zhipu/CogView: fixed WxH grid, URL result downloaded, no edits', async () => {
    const fetchMock = vi.fn(async (url: string) =>
      url.endsWith('/images/generations')
        ? jsonResponse({ data: [{ url: 'https://cdn.example/x.png' }] })
        : new Response(PNG, { status: 200, headers: { 'content-type': 'image/png' } }),
    )
    vi.stubGlobal('fetch', fetchMock)
    const out = await generateImageWithProvider('glm', cfg('cogview-4-250304'), {
      prompt: 'a fox',
      aspectRatio: '9:16',
    })
    expect(out.mime).toBe('image/png')
    const [url, init] = fetchMock.mock.calls[0] as unknown as [string, RequestInit]
    expect(url).toBe('https://open.bigmodel.cn/api/paas/v4/images/generations')
    expect(JSON.parse(init.body as string).size).toBe('768x1344')
    expect((fetchMock.mock.calls[1] as unknown as [string])[0]).toBe('https://cdn.example/x.png')
    await expect(
      generateImageWithProvider('glm', cfg('cogview-4-250304'), {
        prompt: 'x',
        references: [{ bytes: PNG, mime: 'image/png' }],
      }),
    ).rejects.toThrow(/cannot edit/)
  })

  it('xAI: no size field, b64 requested', async () => {
    const fetchMock = vi.fn(async () => jsonResponse({ data: [{ b64_json: PNG_B64 }] }))
    vi.stubGlobal('fetch', fetchMock)
    await generateImageWithProvider('xai', cfg('grok-imagine-image-2.0'), {
      prompt: 'a fox',
      aspectRatio: '16:9',
    })
    const [url, init] = fetchMock.mock.calls[0] as unknown as [string, RequestInit]
    expect(url).toBe('https://api.x.ai/v1/images/generations')
    const body = JSON.parse(init.body as string)
    expect(body.size).toBeUndefined()
    expect(body.response_format).toBe('b64_json')
  })

  it('DashScope/Qwen-Image: multimodal-generation body, W*H size, image URL in output.choices', async () => {
    const fetchMock = vi.fn(async (url: string) =>
      url.includes('multimodal-generation')
        ? jsonResponse({
            output: {
              choices: [{ message: { content: [{ image: 'https://oss.example/q.png' }] } }],
            },
          })
        : new Response(PNG, { status: 200, headers: { 'content-type': 'image/png' } }),
    )
    vi.stubGlobal('fetch', fetchMock)
    await generateImageWithProvider(
      'qwen',
      cfg('qwen-image-plus', 'https://dashscope-intl.aliyuncs.com/compatible-mode/v1'),
      { prompt: 'a fox', aspectRatio: '4:3' },
    )
    const [url, init] = fetchMock.mock.calls[0] as unknown as [string, RequestInit]
    expect(url).toBe(
      'https://dashscope-intl.aliyuncs.com/api/v1/services/aigc/multimodal-generation/generation',
    )
    const body = JSON.parse(init.body as string)
    expect(body.model).toBe('qwen-image-plus')
    expect(body.input.messages[0].content[0].text).toBe('a fox')
    expect(body.parameters).toEqual({ n: 1, watermark: false, size: '1472*1104' })
  })

  it('MiniMax: image_generation with aspect_ratio and base64 result; base_resp errors surface', async () => {
    const fetchMock = vi.fn(async () =>
      jsonResponse({ data: { image_base64: [PNG_B64] }, base_resp: { status_code: 0 } }),
    )
    vi.stubGlobal('fetch', fetchMock)
    const out = await generateImageWithProvider('minimax', cfg('image-01'), {
      prompt: 'a fox',
      aspectRatio: '16:9',
    })
    expect(out.mime).toBe('image/png')
    const [url, init] = fetchMock.mock.calls[0] as unknown as [string, RequestInit]
    expect(url).toBe('https://api.minimax.io/v1/image_generation')
    expect(JSON.parse(init.body as string)).toMatchObject({
      model: 'image-01',
      aspect_ratio: '16:9',
      response_format: 'base64',
    })
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => jsonResponse({ base_resp: { status_code: 1004, status_msg: 'auth' } })),
    )
    await expect(
      generateImageWithProvider('minimax', cfg('image-01'), { prompt: 'x' }),
    ).rejects.toThrow(/1004 auth/)
  })
})

describe('analyzeMediaWithProvider', () => {
  it('sends images as data URLs through chat completions and rejects video for OpenAI-style providers', async () => {
    const fetchMock = vi.fn(async () =>
      jsonResponse({ choices: [{ message: { content: 'two cats' } }] }),
    )
    vi.stubGlobal('fetch', fetchMock)
    const config = { apiKey: 'sk', imageModel: '', analysisModel: 'gpt-5.6-luna' }
    const text = await analyzeMediaWithProvider('openai', config, {
      media: [{ bytes: PNG, mime: 'image/png' }],
      requirements: 'count the cats',
    })
    expect(text).toBe('two cats')
    const body = JSON.parse(
      (fetchMock.mock.calls[0] as unknown as [string, RequestInit])[1].body as string,
    )
    expect(body.model).toBe('gpt-5.6-luna')
    expect(body.messages[0].content[1].image_url.url).toBe(`data:image/png;base64,${PNG_B64}`)
    await expect(
      analyzeMediaWithProvider('openai', config, {
        media: [{ bytes: PNG, mime: 'video/mp4', name: 'clip.mp4' }],
        requirements: 'summarize',
      }),
    ).rejects.toThrow(/video and audio analysis needs/)
    expect(fetchMock).toHaveBeenCalledTimes(1)
  })

  it('sends video as a video_url part for vendors that take it (Qwen via compatible-mode)', async () => {
    const fetchMock = vi.fn(async () =>
      jsonResponse({ choices: [{ message: { content: 'a talk' } }] }),
    )
    vi.stubGlobal('fetch', fetchMock)
    const text = await analyzeMediaWithProvider(
      'qwen',
      { apiKey: 'k', imageModel: '', analysisModel: 'qwen3-vl-plus' },
      { media: [{ bytes: PNG, mime: 'video/mp4' }], requirements: 'summarize' },
    )
    expect(text).toBe('a talk')
    const [url, init] = fetchMock.mock.calls[0] as unknown as [string, RequestInit]
    expect(url).toBe('https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions')
    const body = JSON.parse(init.body as string)
    expect(body.messages[0].content[1]).toEqual({
      type: 'video_url',
      video_url: { url: `data:video/mp4;base64,${PNG_B64}` },
    })
    await expect(
      analyzeMediaWithProvider(
        'qwen',
        { apiKey: 'k', imageModel: '', analysisModel: 'qwen3-vl-plus' },
        { media: [{ bytes: PNG, mime: 'audio/mpeg' }], requirements: 'x' },
      ),
    ).rejects.toThrow(/audio analysis needs/)
  })

  it('inlines small media for Gemini and joins the text parts', async () => {
    const fetchMock = vi.fn(async () =>
      jsonResponse({ candidates: [{ content: { parts: [{ text: 'a ' }, { text: 'talk' }] } }] }),
    )
    vi.stubGlobal('fetch', fetchMock)
    const text = await analyzeMediaWithProvider(
      'gemini',
      { apiKey: 'AIza', imageModel: '', analysisModel: 'gemini-3.7-flash' },
      { media: [{ bytes: PNG, mime: 'video/mp4' }], requirements: 'summarize' },
    )
    expect(text).toBe('a talk')
    const body = JSON.parse(
      (fetchMock.mock.calls[0] as unknown as [string, RequestInit])[1].body as string,
    )
    expect(body.contents[0].parts[0].inline_data.mime_type).toBe('video/mp4')
    expect(body.contents[0].parts[1].text).toBe('summarize')
  })
})

describe('testMediaProvider', () => {
  it('lists models as a credential check and reports failures as text', async () => {
    const fetchMock = vi.fn(async () => jsonResponse({ data: [] }))
    vi.stubGlobal('fetch', fetchMock)
    expect(
      await testMediaProvider('openai', { apiKey: 'sk', imageModel: '', analysisModel: '' }),
    ).toEqual({ ok: true })
    expect((fetchMock.mock.calls[0] as unknown as [string])[0]).toBe(
      'https://api.openai.com/v1/models',
    )
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => jsonResponse({ error: 'nope' }, 403)),
    )
    const failed = await testMediaProvider('gemini', {
      apiKey: 'x',
      imageModel: '',
      analysisModel: '',
    })
    expect(failed.ok).toBe(false)
    expect(failed.error).toMatch(/403/)
  })
})
