import { describe, it, expect, vi, afterEach, beforeAll } from 'vitest'
import { webSearch, imageSearch } from '../src/index'
import { searchOptionsFromSettings, testSearchProvider } from '../src/search-tools'
import { defaultAiSettings } from '@genoffice/ai-provider'

// These cases only test the Serper/DuckDuckGo paths; a local gsk login would take priority, so disable it explicitly
beforeAll(() => {
  process.env.AI_SEARCH_DISABLE_GSK = '1'
})

const realFetch = globalThis.fetch
afterEach(() => {
  globalThis.fetch = realFetch
  delete process.env.SERPER_API_KEY
  delete process.env.TAVILY_API_KEY
})

function mockFetch(
  handler: (url: string, init?: RequestInit) => { ok: boolean; json?: any; text?: string },
) {
  globalThis.fetch = vi.fn(async (url: any, init: any) => {
    const r = handler(String(url), init)
    return {
      ok: r.ok,
      status: r.ok ? 200 : 500,
      headers: new Map(),
      json: async () => r.json,
      text: async () => r.text ?? '',
    } as any
  }) as any
}

describe('webSearch (Serper)', () => {
  it('parses organic results + answer box', async () => {
    process.env.SERPER_API_KEY = 'test-key'
    mockFetch((url) => {
      expect(url).toBe('https://google.serper.dev/search')
      return {
        ok: true,
        json: {
          answerBox: { answer: '42' },
          organic: [
            { title: 'A', link: 'https://a.com', snippet: 'sa' },
            { title: 'B', link: 'https://b.com', snippet: 'sb' },
          ],
        },
      }
    })
    const r = await webSearch('meaning of life', 5)
    expect(r.method).toBe('serper')
    expect(r.answer).toBe('42')
    expect(r.results).toHaveLength(2)
    expect(r.results[0]).toEqual({ title: 'A', url: 'https://a.com', snippet: 'sa' })
  })

  it('falls back to DuckDuckGo when no key', async () => {
    mockFetch((url) => {
      expect(url).toContain('duckduckgo.com')
      return {
        ok: true,
        text: '<a class="result__a" href="/l/?uddg=https%3A%2F%2Fx.com">X Title</a>',
      }
    })
    const r = await webSearch('q', 3)
    expect(r.method).toBe('duckduckgo')
    expect(r.results[0]?.url).toBe('https://x.com')
    expect(r.results[0]?.title).toBe('X Title')
  })
})

describe('webSearch (Tavily)', () => {
  it('parses results + answer when Serper is unconfigured', async () => {
    process.env.TAVILY_API_KEY = 'test-key'
    mockFetch((url, init) => {
      expect(url).toBe('https://api.tavily.com/search')
      expect(JSON.parse(String((init as any)?.body ?? '{}')).api_key).toBe('test-key')
      return {
        ok: true,
        json: {
          answer: '42',
          results: [
            { title: 'A', url: 'https://a.com', content: 'sa' },
            { title: 'B', url: 'https://b.com', content: 'sb' },
          ],
        },
      }
    })
    const r = await webSearch('meaning of life', 5)
    expect(r.method).toBe('tavily')
    expect(r.answer).toBe('42')
    expect(r.results).toHaveLength(2)
    expect(r.results[0]).toEqual({ title: 'A', url: 'https://a.com', snippet: 'sa' })
  })

  it('prefers Serper over Tavily when both keys are set', async () => {
    process.env.SERPER_API_KEY = 'serper-key'
    process.env.TAVILY_API_KEY = 'tavily-key'
    mockFetch((url) => {
      expect(url).toBe('https://google.serper.dev/search')
      return { ok: true, json: { organic: [{ title: 'A', link: 'https://a.com' }] } }
    })
    const r = await webSearch('q', 3)
    expect(r.method).toBe('serper')
  })

  it('falls back to DuckDuckGo when Tavily returns nothing usable', async () => {
    process.env.TAVILY_API_KEY = 'test-key'
    mockFetch((url) => {
      if (url === 'https://api.tavily.com/search') return { ok: true, json: { results: [] } }
      expect(url).toContain('duckduckgo.com')
      return {
        ok: true,
        text: '<a class="result__a" href="/l/?uddg=https%3A%2F%2Fx.com">X Title</a>',
      }
    })
    const r = await webSearch('q', 3)
    expect(r.method).toBe('duckduckgo')
  })
})

describe('DuckDuckGo fallback error surfacing', () => {
  it('web: reports method error when the backend is unreachable', async () => {
    mockFetch(() => {
      throw new Error('network down')
    })
    const r = await webSearch('q', 3)
    expect(r.method).toBe('error')
    expect(r.results).toHaveLength(0)
    expect(r.error).toContain('duckduckgo')
  })

  it('web: stays a plain empty result when the backend responds with nothing', async () => {
    mockFetch(() => ({ ok: true, text: '<html></html>' }))
    const r = await webSearch('q', 3)
    expect(r.method).toBe('duckduckgo')
    expect(r.results).toHaveLength(0)
    expect(r.error).toBeUndefined()
  })

  it('images: reports method error when the backend is unreachable', async () => {
    mockFetch(() => {
      throw new Error('network down')
    })
    const r = await imageSearch('cats', 8)
    expect(r.method).toBe('error')
    expect(r.images).toHaveLength(0)
    expect(r.error).toContain('duckduckgo')
  })
})

describe('imageSearch (Serper)', () => {
  it('parses images + filters copyright hosts', async () => {
    process.env.SERPER_API_KEY = 'test-key'
    mockFetch((url) => {
      expect(url).toBe('https://google.serper.dev/images')
      return {
        ok: true,
        json: {
          images: [
            {
              title: 'good',
              imageUrl: 'https://cdn.example.com/a.jpg',
              link: 'https://example.com',
              imageWidth: 800,
              imageHeight: 600,
            },
            {
              title: 'paid',
              imageUrl: 'https://gettyimages.com/x.jpg',
              link: 'https://gettyimages.com',
            },
          ],
        },
      }
    })
    const r = await imageSearch('cats', 8)
    expect(r.method).toBe('serper')
    expect(r.images).toHaveLength(1) // getty is filtered out
    expect(r.images[0]).toMatchObject({
      imageUrl: 'https://cdn.example.com/a.jpg',
      width: 800,
      height: 600,
    })
  })
})

describe('webSearch (SearchOptions)', () => {
  it('uses a caller-supplied Serper key instead of the env var', async () => {
    const seen: string[] = []
    mockFetch((url, init) => {
      seen.push(String((init?.headers as Record<string, string>)['X-API-KEY']))
      return { ok: true, json: { organic: [{ title: 'A', link: 'https://a.com', snippet: 's' }] } }
    })
    const r = await webSearch('q', 3, { useGsk: false, serperKey: 'user-key' })
    expect(r.method).toBe('serper')
    expect(seen).toEqual(['user-key'])
  })

  it('tries Tavily first when preferred and never touches Serper on success', async () => {
    const urls: string[] = []
    mockFetch((url) => {
      urls.push(url)
      return { ok: true, json: { results: [{ title: 'T', url: 'https://t.com', content: 'c' }] } }
    })
    const r = await webSearch('q', 3, {
      useGsk: false,
      tavilyKey: 'tv',
      serperKey: 'sp',
      prefer: 'tavily',
    })
    expect(r.method).toBe('tavily')
    expect(urls).toEqual(['https://api.tavily.com/search'])
  })
})

describe('search-tools', () => {
  it('maps the settings block onto SearchOptions', () => {
    const base = defaultAiSettings()
    expect(searchOptionsFromSettings(base)).toEqual({ useGsk: true })
    expect(searchOptionsFromSettings({ ...base, gskToolsEnabled: false })).toEqual({
      useGsk: false,
    })
    const serper = {
      ...base,
      search: {
        provider: 'serper' as const,
        providers: { serper: { apiKey: 'k' }, tavily: { apiKey: '' } },
      },
    }
    expect(searchOptionsFromSettings(serper)).toEqual({ useGsk: false, serperKey: 'k' })
    const tavily = {
      ...base,
      search: {
        provider: 'tavily' as const,
        providers: { serper: { apiKey: '' }, tavily: { apiKey: 't' } },
      },
    }
    expect(searchOptionsFromSettings(tavily)).toEqual({
      useGsk: false,
      tavilyKey: 't',
      prefer: 'tavily',
    })
    // no key → genspark chain
    const empty = {
      ...base,
      search: {
        provider: 'serper' as const,
        providers: { serper: { apiKey: '' }, tavily: { apiKey: '' } },
      },
    }
    expect(searchOptionsFromSettings(empty)).toEqual({ useGsk: true })
  })

  it('reports a rejected key as a failure instead of the silent free fallback', async () => {
    mockFetch((url) => {
      if (url.includes('serper')) return { ok: false }
      return { ok: true, text: '<a class="result__a" href="/l/?uddg=https%3A%2F%2Fx.com">X</a>' }
    })
    const bad = await testSearchProvider('serper', 'wrong')
    expect(bad.ok).toBe(false)
    expect(bad.error).toMatch(/serper/)
    mockFetch(() => ({
      ok: true,
      json: { organic: [{ title: 'A', link: 'https://a.com', snippet: 's' }] },
    }))
    expect(await testSearchProvider('serper', 'right')).toEqual({ ok: true })
    expect(await testSearchProvider('tavily', '')).toEqual({ ok: false, error: 'API key is empty' })
  })
})
