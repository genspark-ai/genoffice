import { describe, expect, it } from 'vitest'
import {
  AI_PROVIDERS,
  DEFAULT_MAX_OUTPUT_TOKENS,
  MAX_MAX_OUTPUT_TOKENS,
  MIN_MAX_OUTPUT_TOKENS,
  activeProvider,
  clampMaxOutputTokens,
  cloudToolsEnabled,
  defaultAiSettings,
  maxOutputTokensOf,
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

describe('provider model catalog', () => {
  it('offers DeepSeek Vision Exp only through the direct BYOK provider', () => {
    const genspark = AI_PROVIDERS.find((provider) => provider.id === 'genspark')!
    const deepseek = AI_PROVIDERS.find((provider) => provider.id === 'deepseek')!

    expect(deepseek.models).toContain('deepseek-v4-flash-vision-exp')
    expect(genspark.models).not.toContain('deep-seek-v4-flash')
    expect(genspark.models).not.toContain('deep-seek-v4-flash-vision-exp-openrouter')
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

  it('rewrites a stored model id the vendor has retired', () => {
    const resolved = resolveAiSettings(
      {
        providers: {
          deepseek: { apiKey: 'sk-user', model: 'deepseek-reasoner' },
        } as never,
      },
      defaultAiSettings(),
    )
    expect(resolved.providers.deepseek).toEqual({ apiKey: 'sk-user', model: 'deepseek-v4-flash' })
  })

  it('rewrites genspark model ids the proxy no longer serves', () => {
    const resolved = resolveAiSettings(
      {
        providers: {
          genspark: { apiKey: '', model: 'gemini-3.7-flash' },
        } as never,
      },
      defaultAiSettings(),
    )
    expect(resolved.providers.genspark.model).toBe('claude-opus-4-7')

    const gpt = resolveAiSettings(
      {
        providers: {
          genspark: { apiKey: '', model: 'gpt-5.6' },
        } as never,
      },
      defaultAiSettings(),
    )
    expect(gpt.providers.genspark.model).toBe('gpt-5.6-terra')
  })

  it('leaves a still-supported model id alone', () => {
    const resolved = resolveAiSettings(
      {
        providers: {
          deepseek: { apiKey: 'sk-user', model: 'deepseek-v4-pro' },
        } as never,
      },
      defaultAiSettings(),
    )
    expect(resolved.providers.deepseek.model).toBe('deepseek-v4-pro')
  })

  it('trims whitespace pasted around stored keys and base URLs', () => {
    const resolved = resolveAiSettings(
      {
        providers: {
          deepseek: { apiKey: ' sk-user\n', model: 'deepseek-v4-pro' },
          custom: { apiKey: 'k', model: 'm', baseUrl: ' http://localhost:1234/v1 ' },
        } as never,
      },
      defaultAiSettings(),
    )
    expect(resolved.providers.deepseek.apiKey).toBe('sk-user')
    expect(resolved.providers.deepseek.baseUrl).toBeUndefined()
    expect(resolved.providers.custom.baseUrl).toBe('http://localhost:1234/v1')
  })

  it('trims the legacy single-endpoint key and base URL too', () => {
    const resolved = resolveAiSettings(
      { apiKey: ' legacy-key ', baseUrl: ' https://legacy.example.com/v1 ' },
      defaultAiSettings(),
    )
    expect(resolved.providers.custom.apiKey).toBe('legacy-key')
    expect(resolved.providers.custom.baseUrl).toBe('https://legacy.example.com/v1')
  })

  it('carries a stored output cap and clamps a hand-edited one', () => {
    // a multi-provider file (the legacy single-endpoint shape returns defaults wholesale)
    const stored = { providers: {} as never }
    expect(
      resolveAiSettings({ ...stored, maxOutputTokens: 32768 }, defaultAiSettings()).maxOutputTokens,
    ).toBe(32768)
    // a settings file edited by hand must not forward an absurd budget to the endpoint
    expect(
      resolveAiSettings({ ...stored, maxOutputTokens: 1 }, defaultAiSettings()).maxOutputTokens,
    ).toBe(MIN_MAX_OUTPUT_TOKENS)
    expect(
      resolveAiSettings({ ...stored, maxOutputTokens: 1e9 }, defaultAiSettings()).maxOutputTokens,
    ).toBe(MAX_MAX_OUTPUT_TOKENS)
    // absent stays absent: pre-existing settings files keep the default behaviour
    expect('maxOutputTokens' in resolveAiSettings(stored, defaultAiSettings())).toBe(false)
  })
})

describe('maxOutputTokensOf', () => {
  it('falls back to the default when the setting is absent or unusable', () => {
    expect(maxOutputTokensOf({})).toBe(DEFAULT_MAX_OUTPUT_TOKENS)
    expect(maxOutputTokensOf(undefined)).toBe(DEFAULT_MAX_OUTPUT_TOKENS)
    expect(maxOutputTokensOf({ maxOutputTokens: Number.NaN })).toBe(DEFAULT_MAX_OUTPUT_TOKENS)
    expect(maxOutputTokensOf({ maxOutputTokens: '8192' as unknown as number })).toBe(
      DEFAULT_MAX_OUTPUT_TOKENS,
    )
  })

  it('honors a stored cap inside the bounds', () => {
    expect(maxOutputTokensOf({ maxOutputTokens: 16384 })).toBe(16384)
    expect(maxOutputTokensOf({ maxOutputTokens: 3.7 })).toBe(MIN_MAX_OUTPUT_TOKENS)
    expect(maxOutputTokensOf({ maxOutputTokens: 20000 })).toBe(20000)
  })
})

describe('clampMaxOutputTokens', () => {
  it('floors, bounds and defaults whatever the settings field or the input box held', () => {
    expect(clampMaxOutputTokens(16384.9)).toBe(16384)
    expect(clampMaxOutputTokens(0)).toBe(MIN_MAX_OUTPUT_TOKENS)
    expect(clampMaxOutputTokens(5e6)).toBe(MAX_MAX_OUTPUT_TOKENS)
    expect(clampMaxOutputTokens(Number.POSITIVE_INFINITY)).toBe(DEFAULT_MAX_OUTPUT_TOKENS)
    expect(clampMaxOutputTokens(undefined)).toBe(DEFAULT_MAX_OUTPUT_TOKENS)
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
