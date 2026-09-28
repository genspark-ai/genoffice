import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { defaultAiSettings } from '@genoffice/ai-provider'
import { imageSearch, webSearch } from '../src/index'
import { searchOptionsFromSettings, testSearchProvider } from '../src/search-tools'

const webEndpoint = 'https://api.serply.io/v1/search'
const imageEndpoint = 'https://api.serply.io/v1/image'
const webResponse = () =>
  Response.json({
    results: [
      {
        title: 'GenOffice',
        link: 'https://genoffice.ai/',
        description: 'A free, open-source office suite.',
      },
      { title: 'Not a page', link: 'javascript:void(0)', description: 'dropped' },
      { title: 'Docs', link: 'https://genoffice.ai/docs', description: 'Documentation.' },
    ],
    answers: [],
  })
const image = (title: string, src: string, href: string, domain: string) => ({
  image: { src: 'https://encrypted-tbn0.gstatic.com/thumb', alt: title },
  link: { href, title, domain },
  original_image: { src, width: '330', height: '550', file_format: 'image/jpeg' },
})
const fallback = () =>
  new Response('<a class="result__a" href="https://fallback.example.com">Fallback</a>')

beforeEach(() => {
  vi.stubEnv('AI_SEARCH_DISABLE_GSK', '1')
  for (const key of ['SERPER_API_KEY', 'SERPLY_API_KEY', 'TAVILY_API_KEY', 'PARALLEL_API_KEY'])
    vi.stubEnv(key, '')
})

afterEach(() => {
  vi.unstubAllGlobals()
  vi.unstubAllEnvs()
})

describe('Serply web search', () => {
  it('sends a GET with the key header and maps results[].description', async () => {
    const fetch = vi.fn<typeof globalThis.fetch>(async () => webResponse())
    vi.stubGlobal('fetch', fetch)
    const r = await webSearch('genoffice', 5, {
      useGsk: false,
      serplyKey: 'user-key',
      serperKey: 'serper-key',
      prefer: 'serply',
    })
    expect(r).toEqual({
      results: [
        {
          title: 'GenOffice',
          url: 'https://genoffice.ai/',
          snippet: 'A free, open-source office suite.',
        },
        { title: 'Docs', url: 'https://genoffice.ai/docs', snippet: 'Documentation.' },
      ],
      method: 'serply',
    })
    expect(fetch).toHaveBeenCalledOnce()
    const [url, init] = fetch.mock.calls[0]!
    const parsed = new URL(String(url))
    expect(`${parsed.origin}${parsed.pathname}`).toBe(webEndpoint)
    expect(parsed.searchParams.get('q')).toBe('genoffice')
    expect(parsed.searchParams.get('num')).toBe('5')
    expect(init?.method).toBeUndefined()
    expect(init?.headers).toEqual({ 'X-Api-Key': 'user-key', 'User-Agent': 'genoffice' })
  })

  it('reads SERPLY_API_KEY and caps the mapped results at maxResults', async () => {
    vi.stubEnv('SERPLY_API_KEY', 'environment-key')
    vi.stubGlobal(
      'fetch',
      vi.fn<typeof globalThis.fetch>(async () => webResponse()),
    )
    const r = await webSearch('genoffice', 1)
    expect(r.method).toBe('serply')
    expect(r.results.map((x) => x.url)).toEqual(['https://genoffice.ai/'])
  })

  it('falls through to the next backend when Serply rejects the key', async () => {
    const urls: string[] = []
    vi.stubGlobal(
      'fetch',
      vi.fn<typeof globalThis.fetch>(async (input) => {
        const url = String(input)
        urls.push(url)
        if (url.startsWith(webEndpoint))
          return Response.json({ detail: 'Invalid API key' }, { status: 401 })
        if (url === 'https://search.parallel.ai/mcp') return new Response(null, { status: 503 })
        return fallback()
      }),
    )
    const r = await webSearch('genoffice', 3, {
      useGsk: false,
      serplyKey: 'bad-key',
      prefer: 'serply',
    })
    expect(urls[0]!.startsWith(webEndpoint)).toBe(true)
    expect(r.method).toBe('duckduckgo')
  })
})

describe('Serply image search', () => {
  it('maps image_results, parses string dimensions and filters stock hosts', async () => {
    const fetch = vi.fn<typeof globalThis.fetch>(async () =>
      Response.json({
        image_results: [
          image(
            'Tower',
            'https://upload.example.org/tower.jpg',
            'https://example.org/t',
            'example.org',
          ),
          image(
            'Stock',
            'https://media.gettyimages.com/x.jpg',
            'https://gettyimages.com',
            'gettyimages.com',
          ),
          image('Second', 'https://cdn.example.com/b.jpg', 'https://example.com/b', 'example.com'),
          image('Third', 'https://cdn.example.com/c.jpg', 'https://example.com/c', 'example.com'),
        ],
      }),
    )
    vi.stubGlobal('fetch', fetch)
    const r = await imageSearch('eiffel tower', 2, {
      useGsk: false,
      serplyKey: 'user-key',
      serperKey: 'serper-key',
      prefer: 'serply',
    })
    expect(r.method).toBe('serply')
    expect(r.images).toEqual([
      {
        title: 'Tower',
        imageUrl: 'https://upload.example.org/tower.jpg',
        sourceUrl: 'https://example.org/t',
        source: 'example.org',
        width: 330,
        height: 550,
      },
      {
        title: 'Second',
        imageUrl: 'https://cdn.example.com/b.jpg',
        sourceUrl: 'https://example.com/b',
        source: 'example.com',
        width: 330,
        height: 550,
      },
    ])
    expect(fetch).toHaveBeenCalledOnce()
    const parsed = new URL(String(fetch.mock.calls[0]![0]))
    expect(`${parsed.origin}${parsed.pathname}`).toBe(imageEndpoint)
    expect(parsed.searchParams.get('q')).toBe('eiffel tower')
  })

  it('keeps Serper first for images unless Serply is preferred', async () => {
    const urls: string[] = []
    vi.stubGlobal(
      'fetch',
      vi.fn<typeof globalThis.fetch>(async (input) => {
        const url = String(input)
        urls.push(url)
        return Response.json({
          images: [
            { title: 'S', imageUrl: 'https://cdn.example.com/s.jpg', link: 'https://s.com' },
          ],
        })
      }),
    )
    const r = await imageSearch('cats', 3, { useGsk: false, serplyKey: 'a', serperKey: 'b' })
    expect(r.method).toBe('serper')
    expect(urls).toEqual(['https://google.serper.dev/images'])
  })
})

describe('Serply settings wiring', () => {
  it('maps a selected Serply key onto SearchOptions', () => {
    const base = defaultAiSettings()
    const settings = {
      ...base,
      search: {
        provider: 'serply' as const,
        providers: { ...base.search!.providers, serply: { apiKey: ' k ' } },
      },
    }
    expect(searchOptionsFromSettings(settings)).toEqual({
      useGsk: false,
      serplyKey: 'k',
      prefer: 'serply',
    })
  })

  it('tests a Serply key with one query and reports a rejected key', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn<typeof globalThis.fetch>(async () => webResponse()),
    )
    expect(await testSearchProvider('serply', 'right')).toEqual({ ok: true })
    vi.stubGlobal(
      'fetch',
      vi.fn<typeof globalThis.fetch>(async (input) =>
        String(input).startsWith(webEndpoint)
          ? Response.json({ detail: 'Invalid API key' }, { status: 401 })
          : fallback(),
      ),
    )
    const bad = await testSearchProvider('serply', 'wrong')
    expect(bad.ok).toBe(false)
    expect(bad.error).toMatch(/serply/)
    expect(await testSearchProvider('serply', '')).toEqual({ ok: false, error: 'API key is empty' })
  })
})
