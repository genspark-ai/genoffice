import { afterEach, describe, expect, it, vi } from 'vitest'
import { defaultAiMediaSettings, getMediaProviderMeta } from '../src/media'
import { analyzeMediaWithProvider } from '../src/media-protocols'
import { AI_PROVIDER_ADAPTERS, modelLacksVision } from '../src/registry'

afterEach(() => {
  vi.unstubAllGlobals()
})

describe('MiniMax input capabilities', () => {
  it('supports M3 vision while excluding M2.7, including qualified model ids', () => {
    expect(AI_PROVIDER_ADAPTERS.minimax.capabilities.vision).toBe(true)
    for (const model of [
      'MiniMax-M2.7',
      'minimax-m2.7',
      'minimax/MiniMax-M2.7',
      'MiniMax-M2.7-6bit',
    ]) {
      expect(modelLacksVision(model)).toBe(true)
    }
    expect(modelLacksVision('MiniMax-M3')).toBe(false)
    expect(modelLacksVision('minimax/MiniMax-M3')).toBe(false)
  })

  it('defaults media analysis to M3 without offering the text-only M2.7 model', () => {
    const meta = getMediaProviderMeta('minimax')!
    expect(meta.analysisModels).toEqual(['MiniMax-M3'])
    expect(meta.videoAnalysis).toBe(true)
    expect(defaultAiMediaSettings().providers.minimax.analysisModel).toBe('MiniMax-M3')
  })

  it.each([undefined, 'https://api.minimaxi.com/v1'])(
    'sends images and video through the configured MiniMax endpoint (%s)',
    async (baseUrl) => {
      const fetchMock = vi.fn(
        async () =>
          new Response(
            JSON.stringify({ choices: [{ message: { content: 'A short clip and an image.' } }] }),
            { status: 200, headers: { 'Content-Type': 'application/json' } },
          ),
      )
      vi.stubGlobal('fetch', fetchMock)
      const config = { apiKey: 'test-key', imageModel: '', analysisModel: '', baseUrl }
      const media = [
        { bytes: new Uint8Array([1, 2, 3]), mime: 'image/png' },
        { bytes: new Uint8Array([4, 5, 6]), mime: 'video/mp4' },
      ]
      expect(
        await analyzeMediaWithProvider('minimax', config, {
          media,
          requirements: 'Describe the media',
        }),
      ).toBe('A short clip and an image.')
      const [url, init] = fetchMock.mock.calls[0] as unknown as [string, RequestInit]
      expect(url).toBe(`${baseUrl || 'https://api.minimax.io/v1'}/chat/completions`)
      expect(init.headers).toMatchObject({ Authorization: 'Bearer test-key' })
      expect(JSON.parse(init.body as string)).toEqual({
        model: 'MiniMax-M3',
        messages: [
          {
            role: 'user',
            content: [
              { type: 'text', text: 'Describe the media' },
              { type: 'image_url', image_url: { url: 'data:image/png;base64,AQID' } },
              { type: 'video_url', video_url: { url: 'data:video/mp4;base64,BAUG' } },
            ],
          },
        ],
      })
    },
  )
})
