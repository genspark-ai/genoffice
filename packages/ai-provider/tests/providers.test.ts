import { describe, expect, it } from 'vitest'
import {
  AI_PROVIDERS,
  activeProvider,
  cloudToolsEnabled,
  defaultAiSettings,
  resolveAiSettings,
} from '../src/providers'
import type { AiProviderId } from '../src/types'

describe('defaultAiSettings', () => {
  it('gives every provider its default model and an empty key by default', () => {
    const settings = defaultAiSettings()
    expect(settings.provider).toBe('genspark')
    for (const meta of AI_PROVIDERS) {
      expect(settings.providers[meta.id].apiKey).toBe('')
      expect(settings.providers[meta.id].model).toBe(meta.defaultModel)
    }
    expect(settings.providers.custom.baseUrl).toBe('')
    expect(settings.providers.anthropic.baseUrl).toBeUndefined()
  })

  it('applies caller-supplied default keys only to the listed providers', () => {
    const settings = defaultAiSettings({ anthropic: 'sk-ant-preset' })
    expect(settings.providers.anthropic.apiKey).toBe('sk-ant-preset')
    expect(settings.providers.gemini.apiKey).toBe('')
  })
})

describe('resolveAiSettings', () => {
  it('returns fresh defaults when nothing is stored', () => {
    const defaults = defaultAiSettings({ anthropic: 'sk-ant-preset' })
    expect(resolveAiSettings({}, defaults)).toEqual(defaults)
  })

  it('migrates the pre-provider single-endpoint shape into the custom provider', () => {
    const defaults = defaultAiSettings()
    const resolved = resolveAiSettings(
      { apiKey: 'legacy-key', model: 'legacy-model', baseUrl: 'https://legacy.example.com/v1' },
      defaults,
    )
    expect(resolved.providers.custom).toEqual({
      apiKey: 'legacy-key',
      model: 'legacy-model',
      baseUrl: 'https://legacy.example.com/v1',
    })
    // untouched providers keep their defaults
    expect(resolved.providers.anthropic).toEqual(defaults.providers.anthropic)
  })

  it('defaults the legacy base URL to the OpenAI endpoint when omitted', () => {
    const resolved = resolveAiSettings({ apiKey: 'legacy-key' }, defaultAiSettings())
    expect(resolved.providers.custom.baseUrl).toBe('https://api.openai.com/v1')
  })

  it('merges stored multi-provider settings over the defaults, provider by provider', () => {
    const defaults = defaultAiSettings({ anthropic: 'preset-key' })
    const resolved = resolveAiSettings(
      {
        provider: 'gemini',
        providers: {
          gemini: { apiKey: 'stored-gemini-key', model: 'gemini-2.5-pro' },
        } as never,
      },
      defaults,
    )
    expect(resolved.provider).toBe('gemini')
    expect(resolved.providers.gemini).toEqual({
      apiKey: 'stored-gemini-key',
      model: 'gemini-2.5-pro',
    })
    // provider not mentioned in stored.providers keeps the computed default
    expect(resolved.providers.anthropic.apiKey).toBe('preset-key')
  })
})

describe('activeProvider', () => {
  it('honors a configured BYOK provider and falls back to genspark otherwise', () => {
    const settings = defaultAiSettings()
    expect(activeProvider(settings)).toBe('genspark')

    settings.provider = 'kimi'
    expect(activeProvider(settings)).toBe('genspark') // no key yet
    settings.providers.kimi.apiKey = 'sk-user'
    expect(activeProvider(settings)).toBe('kimi')
  })

  it('requires a base URL for providers that declare needsBaseUrl', () => {
    const settings = defaultAiSettings()
    settings.provider = 'custom'
    settings.providers.custom.apiKey = 'k'
    expect(activeProvider(settings)).toBe('genspark')
    settings.providers.custom.baseUrl = 'http://localhost:1234/v1'
    expect(activeProvider(settings)).toBe('genspark') // custom's default model is empty
    settings.providers.custom.model = 'my-model'
    expect(activeProvider(settings)).toBe('custom')
  })

  it('falls back to genspark for unknown ids from a hand-edited settings file', () => {
    const settings = defaultAiSettings()
    settings.provider = 'nonsense' as AiProviderId
    expect(activeProvider(settings)).toBe('genspark')
  })

  it('genspark never requires a key (injected from the gsk login at request time)', () => {
    const settings = defaultAiSettings()
    settings.provider = 'genspark'
    expect(activeProvider(settings)).toBe('genspark')
  })
})

describe('gskToolsEnabled', () => {
  it('defaults on, survives resolveAiSettings, and only an explicit false turns it off', () => {
    expect(cloudToolsEnabled(defaultAiSettings())).toBe(true)
    // pre-toggle settings file (field absent) stays on
    const legacy = resolveAiSettings({ providers: {} as never }, defaultAiSettings())
    expect(cloudToolsEnabled(legacy)).toBe(true)
    const off = resolveAiSettings(
      { providers: {} as never, gskToolsEnabled: false },
      defaultAiSettings(),
    )
    expect(off.gskToolsEnabled).toBe(false)
    expect(cloudToolsEnabled(off)).toBe(false)
  })
})
