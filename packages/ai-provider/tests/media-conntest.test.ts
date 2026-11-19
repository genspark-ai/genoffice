import { afterEach, describe, expect, it, vi } from 'vitest'
import { testMediaProvider } from '../src/media-protocols'

afterEach(() => {
  vi.unstubAllGlobals()
})

function errorResponse(body: unknown, status: number): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json' },
  })
}

const config = { apiKey: 'sk-test', imageModel: '', analysisModel: '' }

describe('testMediaProvider connection test', () => {
  it('treats 404 as ok (vendor without a model-listing endpoint)', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => errorResponse({ error: 'not found' }, 404)),
    )
    const result = await testMediaProvider('openai', config)
    expect(result).toEqual({ ok: true })
  })

  it('treats 405 as ok (vendor without a model-listing endpoint)', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => errorResponse({ error: 'method not allowed' }, 405)),
    )
    const result = await testMediaProvider('openai', config)
    expect(result).toEqual({ ok: true })
  })

  it('surfaces 429 rate limits as ok:false with status and rate-limit wording', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => errorResponse({ error: { message: 'too many requests' } }, 429)),
    )
    const result = await testMediaProvider('openai', config)
    expect(result.ok).toBe(false)
    expect(result.error).toMatch(/429/)
    expect(result.error).toMatch(/rate limit/i)
  })

  it('surfaces 500 server errors as ok:false with status and server-error wording', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => errorResponse({ error: { message: 'internal error' } }, 500)),
    )
    const result = await testMediaProvider('openai', config)
    expect(result.ok).toBe(false)
    expect(result.error).toMatch(/500/)
    expect(result.error).toMatch(/server error/i)
  })

  it('appends /models on the path of an OpenAI-shaped base that carries a query', async () => {
    const fetchMock = vi.fn(async () => errorResponse({ data: [] }, 404))
    vi.stubGlobal('fetch', fetchMock)
    await testMediaProvider('custom', {
      ...config,
      baseUrl: 'https://gw.example.com/v1?key=abc',
      imageModel: '',
      analysisModel: '',
    })
    expect((fetchMock.mock.calls[0] as unknown as [string])[0]).toBe(
      'https://gw.example.com/v1/models?key=abc',
    )
  })

  it('merges pageSize into the query of a Gemini base that already carries one', async () => {
    const fetchMock = vi.fn(async () => errorResponse({ models: [] }, 200))
    vi.stubGlobal('fetch', fetchMock)
    const result = await testMediaProvider('gemini', {
      ...config,
      baseUrl: 'https://gw.example.com/v1beta?key=abc',
      imageModel: '',
      analysisModel: '',
    })
    expect(result).toEqual({ ok: true })
    expect((fetchMock.mock.calls[0] as unknown as [string])[0]).toBe(
      'https://gw.example.com/v1beta/models?key=abc&pageSize=1',
    )
  })
})
