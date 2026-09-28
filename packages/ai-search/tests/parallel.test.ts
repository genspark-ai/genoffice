import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { defaultAiSettings } from '@genoffice/ai-provider'
import { imageSearch, webSearch } from '../src/index'
import { searchOptionsFromSettings, testSearchProvider, webSearchTool } from '../src/search-tools'

const endpoint = 'https://api.parallel.ai/v1/search'
const result = { title: 'GenOffice', url: 'https://example.com', excerpts: ['First.', 'Second.'] }
const response = () =>
  Response.json({ search_id: 'test-search', session_id: 'test-session', results: [result] })
const fallback = () =>
  new Response('<a class="result__a" href="https://fallback.example.com">Fallback</a>')
let dir: string

beforeEach(() => {
  vi.stubEnv('AI_SEARCH_DISABLE_GSK', '1')
  for (const key of ['SERPER_API_KEY', 'SERPLY_API_KEY', 'TAVILY_API_KEY', 'PARALLEL_API_KEY'])
    vi.stubEnv(key, '')
  dir = mkdtempSync(join(tmpdir(), 'genoffice-parallel-'))
})

afterEach(() => {
  vi.unstubAllGlobals()
  vi.unstubAllEnvs()
  vi.useRealTimers()
  rmSync(dir, { recursive: true, force: true })
})

describe('Parallel search', () => {
  it('uses the selected key first, sends the v1 request, and limits mapped excerpts', async () => {
    vi.stubEnv('PARALLEL_API_KEY', 'environment-key')
    vi.stubEnv('SERPER_API_KEY', 'serper-key')
    vi.stubEnv('TAVILY_API_KEY', 'tavily-key')
    const fetch = vi.fn<typeof globalThis.fetch>(async () =>
      Response.json({ results: [result, { ...result, title: 'Extra' }] }),
    )
    vi.stubGlobal('fetch', fetch)
    expect(
      await webSearch('office tools', 1, {
        useGsk: false,
        parallelKey: 'saved-key',
        prefer: 'parallel',
      }),
    ).toEqual({
      method: 'parallel',
      results: [{ title: 'GenOffice', url: 'https://example.com', snippet: 'First.\nSecond.' }],
    })
    expect(fetch).toHaveBeenCalledTimes(1)
    const [url, init] = fetch.mock.calls[0] as unknown as [string, RequestInit]
    expect(url).toBe(endpoint)
    expect(init.method).toBe('POST')
    expect(new Headers(init.headers).get('x-api-key')).toBe('saved-key')
    expect(JSON.parse(String(init.body))).toEqual({
      search_queries: ['office tools'],
      mode: 'fast',
    })
    expect(init.signal).toBeInstanceOf(AbortSignal)
  })

  it('uses the environment key when no explicit key is supplied', async () => {
    vi.stubEnv('PARALLEL_API_KEY', 'environment-key')
    const fetch = vi.fn<typeof globalThis.fetch>(async () => response())
    vi.stubGlobal('fetch', fetch)
    expect((await webSearch('office')).method).toBe('parallel')
    const [url, init] = fetch.mock.calls[0] as unknown as [string, RequestInit]
    expect(url).toBe(endpoint)
    expect(new Headers(init.headers).get('x-api-key')).toBe('environment-key')
  })

  it('keeps existing keyed providers ahead of Parallel unless selected', async () => {
    vi.stubEnv('PARALLEL_API_KEY', 'parallel-key')
    vi.stubEnv('SERPER_API_KEY', 'serper-key')
    const fetch = vi.fn<typeof globalThis.fetch>(async () =>
      Response.json({ organic: [{ title: 'Existing', link: 'https://existing.example.com' }] }),
    )
    vi.stubGlobal('fetch', fetch)
    expect((await webSearch('office')).method).toBe('serper')
    expect(fetch.mock.calls[0]?.[0]).toBe('https://google.serper.dev/search')
  })

  it.each([401, 429, 500])(
    'falls back on HTTP %i without reporting Parallel success',
    async (status) => {
      const urls: string[] = []
      vi.stubGlobal(
        'fetch',
        vi.fn(async (url: string) => {
          urls.push(url)
          return url === endpoint ? new Response('', { status }) : fallback()
        }),
      )
      const r = await webSearch('office', 1, {
        useGsk: false,
        parallelKey: 'key',
        prefer: 'parallel',
      })
      expect(r.method).toBe('duckduckgo')
      expect(r.results[0]?.url).toBe('https://fallback.example.com')
      expect(urls).toHaveLength(2)
      expect(urls[0]).toBe(endpoint)
    },
  )

  it.each([
    {},
    { results: [] },
    { results: [{ title: 'No URL' }, { url: 'javascript:alert(1)' }] },
  ])('falls back on unusable results: %j', async (data) => {
    const fetch = vi.fn(async (url: string) =>
      url === endpoint ? Response.json(data) : fallback(),
    )
    vi.stubGlobal('fetch', fetch)
    expect((await webSearch('office', 1, { parallelKey: 'key', prefer: 'parallel' })).method).toBe(
      'duckduckgo',
    )
    expect(fetch.mock.calls[0]?.[0]).toBe(endpoint)
  })

  it('skips invalid URLs before limiting results and handles absent titles/excerpts', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () =>
        Response.json({
          results: [
            { url: '' },
            { url: 'https://example.com', title: null, excerpts: [null, 'Useful.', 42] },
            result,
          ],
        }),
      ),
    )
    expect((await webSearch('office', 1, { parallelKey: 'key' })).results).toEqual([
      { title: 'https://example.com', url: 'https://example.com', snippet: 'Useful.' },
    ])
  })

  it('stops mapping results once the caller limit is filled', async () => {
    const mapped = new Set<number>()
    const many = Array.from({ length: 40 }, (_, i) => ({
      title: `Result ${i}`,
      url: `https://example.com/${i}`,
      // A getter so the test can see which results were mapped; Response.json would
      // have flattened this into a plain value before the search ever saw it.
      get excerpts() {
        mapped.add(i)
        return [`Excerpt ${i}.`]
      },
    }))
    vi.stubGlobal(
      'fetch',
      vi.fn(
        async () => ({ ok: true, json: async () => ({ results: many }) }) as unknown as Response,
      ),
    )
    const found = await webSearch('office', 3, { parallelKey: 'key', prefer: 'parallel' })
    expect(found.method).toBe('parallel')
    expect(found.results).toHaveLength(3)
    expect(found.results[0]?.snippet).toBe('Excerpt 0.')
    // v1 search returns everything it has, so the 37 results past the limit used to
    // have their excerpts concatenated and then thrown away
    expect([...mapped]).toEqual([0, 1, 2])
  })

  it('aborts a stalled request and continues to the fallback', async () => {
    vi.useFakeTimers()
    let signal: AbortSignal | null | undefined
    vi.stubGlobal(
      'fetch',
      vi.fn((url: string, init: RequestInit) => {
        if (url === endpoint) signal = init.signal
        return url === endpoint
          ? new Promise<Response>((_resolve, reject) =>
              init.signal!.addEventListener('abort', () => reject(new Error('aborted')), {
                once: true,
              }),
            )
          : Promise.resolve(fallback())
      }),
    )
    const pending = webSearch('office', 1, { parallelKey: 'key' })
    await vi.advanceTimersByTimeAsync(15000)
    expect((await pending).method).toBe('duckduckgo')
    expect(signal?.aborted).toBe(true)
  })

  it('loads a saved Parallel key for editor and CLI searches', async () => {
    const path = join(dir, 'ai-settings.json')
    writeFileSync(
      path,
      JSON.stringify({
        ...defaultAiSettings(),
        search: { provider: 'parallel', providers: { parallel: { apiKey: ' saved-key ' } } },
      }),
    )
    const fetch = vi.fn<typeof globalThis.fetch>(async () => response())
    vi.stubGlobal('fetch', fetch)
    expect((await webSearchTool(path, 'office', 1)).method).toBe('parallel')
    const [, init] = fetch.mock.calls[0] as unknown as [string, RequestInit]
    expect(new Headers(init.headers).get('x-api-key')).toBe('saved-key')
  })

  it('tests only the entered key and never mistakes another provider for success', async () => {
    vi.stubEnv('SERPER_API_KEY', 'serper-key')
    vi.stubEnv('TAVILY_API_KEY', 'tavily-key')
    vi.stubEnv('PARALLEL_API_KEY', 'environment-key')
    const urls: string[] = []
    vi.stubGlobal(
      'fetch',
      vi.fn(async (url: string) => {
        urls.push(url)
        return url === endpoint ? new Response('', { status: 401 }) : fallback()
      }),
    )
    expect((await testSearchProvider('parallel', 'invalid-key')).ok).toBe(false)
    expect(urls[0]).toBe(endpoint)
    expect(urls).toHaveLength(2)
    expect(urls[1]).toContain('duckduckgo.com')
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => response()),
    )
    expect(await testSearchProvider('parallel', 'valid-key')).toEqual({ ok: true })
  })

  it('preserves image fallback when Parallel is selected', async () => {
    vi.stubEnv('SERPER_API_KEY', 'image-key')
    const settings = defaultAiSettings()
    settings.search = {
      provider: 'parallel',
      providers: { ...settings.search!.providers, parallel: { apiKey: 'parallel-key' } },
    }
    const fetch = vi.fn<typeof globalThis.fetch>(async () =>
      Response.json({
        images: [
          {
            title: 'Photo',
            imageUrl: 'https://example.com/photo.jpg',
            link: 'https://example.com',
          },
        ],
      }),
    )
    vi.stubGlobal('fetch', fetch)
    expect((await imageSearch('office', 1, searchOptionsFromSettings(settings))).method).toBe(
      'serper',
    )
    expect(fetch.mock.calls[0]?.[0]).toBe('https://google.serper.dev/images')
  })
})
