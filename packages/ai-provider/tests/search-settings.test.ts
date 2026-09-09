import { describe, expect, it } from 'vitest'
import { defaultAiSettings, resolveAiSettings } from '../src/providers'
import {
  activeSearchProvider,
  defaultAiSearchSettings,
  resolveAiSearchSettings,
} from '../src/search-settings'

describe('search settings', () => {
  it('defaults to genspark with empty keys and rides along in defaultAiSettings', () => {
    expect(defaultAiSearchSettings()).toEqual({
      provider: 'genspark',
      providers: { serper: { apiKey: '' }, tavily: { apiKey: '' } },
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
          providers: { serper: { apiKey: '' }, tavily: { apiKey: '' } },
        },
      }),
    ).toBe('genspark')
    expect(
      activeSearchProvider({
        search: {
          provider: 'serper',
          providers: { serper: { apiKey: 'k' }, tavily: { apiKey: '' } },
        },
      }),
    ).toBe('serper')
    expect(activeSearchProvider({ search: { provider: 'bing', providers: {} } as never })).toBe(
      'genspark',
    )
  })
})
