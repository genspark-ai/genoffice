import { describe, expect, it } from 'vitest'
import {
  AI_PROVIDERS,
  OLLAMA_DEFAULT_BASE_URL,
  defaultAiSettings,
  providerRequiresApiKey,
  resolveAiSettings,
} from '../src/providers'

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

describe('ollama provider registration', () => {
  it('AI_PROVIDERS contains ollama with the correct metadata', () => {
    const ollama = AI_PROVIDERS.find((p) => p.id === 'ollama')
    expect(ollama).toBeDefined()
    expect(ollama!.label).toBe('Ollama')
    expect(ollama!.defaultBaseUrl).toBe('http://localhost:11434/v1')
    expect(ollama!.needsBaseUrl).toBe(true)
    // default model is empty — never hard-code an installed model
    expect(ollama!.defaultModel).toBe('')
    expect(ollama!.models).toEqual([])
  })

  it('defaultAiSettings gives ollama the default base URL, an empty key, and an empty model', () => {
    const settings = defaultAiSettings()
    expect(settings.providers.ollama.baseUrl).toBe(OLLAMA_DEFAULT_BASE_URL)
    expect(settings.providers.ollama.apiKey).toBe('')
    expect(settings.providers.ollama.model).toBe('')
  })

  it('resolveAiSettings merges ollama from defaults when not stored', () => {
    const defaults = defaultAiSettings()
    const resolved = resolveAiSettings({ provider: 'anthropic', providers: {} as never }, defaults)
    expect(resolved.providers.ollama.baseUrl).toBe(OLLAMA_DEFAULT_BASE_URL)
    expect(resolved.providers.ollama.apiKey).toBe('')
    expect(resolved.providers.ollama.model).toBe('')
  })
})

describe('providerRequiresApiKey', () => {
  it('returns false for ollama', () => {
    expect(providerRequiresApiKey('ollama')).toBe(false)
  })

  it.each(['genspark', 'anthropic', 'gemini', 'deepseek', 'openai', 'custom'] as const)(
    'returns true for %s',
    (provider) => {
      expect(providerRequiresApiKey(provider)).toBe(true)
    },
  )

  it('returns true for an unknown provider (fail-closed)', () => {
    expect(providerRequiresApiKey('unknown_provider' as never)).toBe(true)
  })
})
