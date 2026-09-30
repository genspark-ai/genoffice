import { describe, expect, it } from 'vitest'
import { defaultAiSettings, resolveAiSettings } from '../src/providers'
import {
  AI_SEARCH_PROVIDERS,
  activeSearchProvider,
  defaultAiSearchSettings,
  resolveAiSearchSettings,
} from '../src/search-settings'

describe('search settings', () => {
  it('defaults to genspark with empty keys and rides along in defaultAiSettings', () => {
    expect(defaultAiSearchSettings()).toEqual({
      provider: 'genspark',
      providers: {
        serper: { apiKey: '' },
        serply: { apiKey: '' },
        tavily: { apiKey: '' },
        parallel: { apiKey: '' },
        exa: { apiKey: '' },
        firecrawl: { apiKey: '' },
      },
    })
    expect(defaultAiSettings().search?.provider).toBe('genspark')
    const resolved = resolveAiSettings(
      { provider: 'genspark', providers: {} as never },
      defaultAiSettings(),
    )
    expect(resolved.search).toEqual(defaultAiSearchSettings())
  })

  it('merges and trims stored keys', () => {
    const s = resolveAiSearchSettings({
      provider: 'tavily',
      providers: { tavily: { apiKey: ' tvly-1 ' } } as never,
    })
    expect(s.provider).toBe('tavily')
    expect(s.providers.tavily.apiKey).toBe('tvly-1')
    expect(s.providers.serper.apiKey).toBe('')
  })

  it('activates a BYOK search provider only with a key', () => {
    expect(activeSearchProvider({ search: undefined })).toBe('genspark')
    expect(
      activeSearchProvider({
        search: {
          provider: 'serper',
          providers: {
            serper: { apiKey: '' },
            serply: { apiKey: '' },
            tavily: { apiKey: '' },
            parallel: { apiKey: '' },
            exa: { apiKey: '' },
            firecrawl: { apiKey: '' },
          },
        },
      }),
    ).toBe('genspark')
    expect(
      activeSearchProvider({
        search: {
          provider: 'serper',
          providers: {
            serper: { apiKey: 'k' },
            serply: { apiKey: '' },
            tavily: { apiKey: '' },
            parallel: { apiKey: '' },
            exa: { apiKey: '' },
            firecrawl: { apiKey: '' },
          },
        },
      }),
    ).toBe('serper')
    expect(
      activeSearchProvider({
        search: {
          provider: 'serper',
          providers: {
            serper: { apiKey: '   ' },
            serply: { apiKey: '' },
            tavily: { apiKey: '' },
            parallel: { apiKey: '' },
            exa: { apiKey: '' },
            firecrawl: { apiKey: '' },
          },
        },
      }),
    ).toBe('genspark')
    expect(activeSearchProvider({ search: { provider: 'bing', providers: {} } as never })).toBe(
      'genspark',
    )
  })
})

describe('Serply search settings', () => {
  it('restores and trims a saved Serply key and activates it only with a key', () => {
    const settings = resolveAiSearchSettings({
      provider: 'serply',
      providers: { serply: { apiKey: ' serply-key ' } } as never,
    })
    expect(settings.provider).toBe('serply')
    expect(settings.providers.serply.apiKey).toBe('serply-key')
    expect(activeSearchProvider({ search: settings })).toBe('serply')
    expect(
      activeSearchProvider({
        search: { ...settings, providers: { ...settings.providers, serply: { apiKey: '  ' } } },
      }),
    ).toBe('genspark')
  })
})

describe('Parallel search settings', () => {
  it('restores and trims a saved Parallel key', () => {
    const settings = resolveAiSearchSettings(
      JSON.parse(
        JSON.stringify({
          provider: 'parallel',
          providers: { parallel: { apiKey: ' parallel-key ' } },
        }),
      ),
    )
    expect(settings.providers.parallel.apiKey).toBe('parallel-key')
    expect(activeSearchProvider({ search: settings })).toBe('parallel')
  })

  it('loads older settings without changing the selected provider or existing keys', () => {
    const settings = resolveAiSearchSettings(
      JSON.parse(
        JSON.stringify({
          provider: 'tavily',
          providers: { tavily: { apiKey: 'existing-key' } },
        }),
      ),
    )
    expect(activeSearchProvider({ search: settings })).toBe('tavily')
    expect(settings.providers.tavily.apiKey).toBe('existing-key')
    expect(settings.providers.parallel.apiKey).toBe('')
  })

  it('keeps Parallel selected for free search when the key is blank', () => {
    const settings = resolveAiSearchSettings(
      JSON.parse(
        JSON.stringify({
          provider: 'parallel',
          providers: { parallel: { apiKey: '   ' } },
        }),
      ),
    )
    expect(activeSearchProvider({ search: settings })).toBe('parallel')
  })
})

describe('exa and firecrawl providers', () => {
  it('are in the catalog without image search and default to empty keys', () => {
    const exa = AI_SEARCH_PROVIDERS.find((m) => m.id === 'exa')
    const firecrawl = AI_SEARCH_PROVIDERS.find((m) => m.id === 'firecrawl')
    expect(exa?.imageSearch).toBe(false)
    expect(firecrawl?.imageSearch).toBe(false)
    const s = defaultAiSearchSettings()
    expect(s.providers.exa.apiKey).toBe('')
    expect(s.providers.firecrawl.apiKey).toBe('')
  })

  it('activate with a key and fall back to genspark without one', () => {
    const mk = (id: 'exa' | 'firecrawl', apiKey: string) => ({
      search: {
        provider: id,
        providers: {
          serper: { apiKey: '' },
          serply: { apiKey: '' },
          tavily: { apiKey: '' },
          parallel: { apiKey: '' },
          exa: { apiKey: id === 'exa' ? apiKey : '' },
          firecrawl: { apiKey: id === 'firecrawl' ? apiKey : '' },
        },
      },
    })
    expect(activeSearchProvider(mk('exa', ''))).toBe('genspark')
    expect(activeSearchProvider(mk('exa', 'exa-1'))).toBe('exa')
    expect(activeSearchProvider(mk('firecrawl', ' '))).toBe('genspark')
    expect(activeSearchProvider(mk('firecrawl', 'fc-1'))).toBe('firecrawl')
  })

  it('resolveAiSearchSettings keeps stored exa/firecrawl keys trimmed', () => {
    const s = resolveAiSearchSettings({
      provider: 'exa',
      providers: {
        exa: { apiKey: ' k ' },
        firecrawl: { apiKey: 'fc ' },
      } as never,
    })
    expect(s.providers.exa.apiKey).toBe('k')
    expect(s.providers.firecrawl.apiKey).toBe('fc')
    expect(s.provider).toBe('exa')
  })
})
