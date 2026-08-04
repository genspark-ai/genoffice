import { describe, expect, it, afterEach } from 'vitest'
import {
  AI_PROVIDERS,
  applyProviderOverrides,
  defaultAiSettings,
  resolveAiSettings,
} from '../src/providers'

afterEach(() => {
  for (const key of [
    'GENOFFICE_AI_PROVIDER',
    'GENOFFICE_AI_BASE_URL',
    'GENOFFICE_AI_MODEL',
    'GENOFFICE_AI_API_KEY',
  ]) {
    delete process.env[key]
  }
})

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
    expect(resolved.providers.gemini).toEqual({ apiKey: 'stored-gemini-key', model: 'gemini-2.5-pro' })
    // provider not mentioned in stored.providers keeps the computed default
    expect(resolved.providers.anthropic.apiKey).toBe('preset-key')
  })
})

describe('applyProviderOverrides', () => {
  it('leaves genspark untouched when no env is configured', () => {
    const settings = defaultAiSettings()
    expect(applyProviderOverrides(settings).provider).toBe('genspark')
    expect(settings.providers.custom.baseUrl).toBe('')
  })

  it('switches to a custom local endpoint from a base URL + model + key', () => {
    process.env.GENOFFICE_AI_BASE_URL = 'http://localhost:11434/v1'
    process.env.GENOFFICE_AI_MODEL = 'qwen2.5:14b'
    process.env.GENOFFICE_AI_API_KEY = 'not-needed'
    const settings = applyProviderOverrides(defaultAiSettings())
    expect(settings.provider).toBe('custom')
    expect(settings.providers.custom.baseUrl).toBe('http://localhost:11434/v1')
    expect(settings.providers.custom.model).toBe('qwen2.5:14b')
    expect(settings.providers.custom.apiKey).toBe('not-needed')
  })

  it('honors an explicit provider while still applying base URL + model', () => {
    process.env.GENOFFICE_AI_PROVIDER = 'openai'
    process.env.GENOFFICE_AI_MODEL = 'gpt-4o'
    process.env.GENOFFICE_AI_API_KEY = 'sk-test'
    const settings = applyProviderOverrides(defaultAiSettings())
    expect(settings.provider).toBe('openai')
    expect(settings.providers.openai.model).toBe('gpt-4o')
    expect(settings.providers.openai.apiKey).toBe('sk-test')
  })

  it('does not mutate the endpoint when provider is non-custom', () => {
    process.env.GENOFFICE_AI_PROVIDER = 'openai'
    process.env.GENOFFICE_AI_BASE_URL = 'http://localhost:11434/v1'
    const settings = applyProviderOverrides(defaultAiSettings())
    expect(settings.provider).toBe('openai')
    // baseUrl is only meaningful on the custom provider
    expect(settings.providers.openai.baseUrl).toBeUndefined()
  })
})
