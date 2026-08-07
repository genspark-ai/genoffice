import { describe, expect, it } from 'vitest'
import {
  AI_PROVIDERS,
  activeProvider,
  applyModelSettings,
  defaultAiSettings,
  isCustomConfigured,
  resolveAiSettings,
  toModelSettings,
} from '../src/providers'
import {
  DEFAULT_TEMPERATURE,
  isReasoningEffort,
  maxTokensField,
  reasoningEffortField,
  resolveMaxTokens,
  temperatureField,
} from '../src/tuning'
import type { AiModelSettings, AiProviderConfig, AiSettings } from '../src/types'

/** settings selecting a fully configured custom endpoint */
function withCustom(custom: Partial<AiSettings['providers']['custom']>): AiSettings {
  const settings = defaultAiSettings()
  settings.provider = 'custom'
  settings.providers.custom = {
    apiKey: '',
    model: 'local-model',
    baseUrl: 'http://localhost:11434/v1',
    ...custom,
  }
  return settings
}

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

describe('isCustomConfigured', () => {
  it('needs both an endpoint and a model name', () => {
    expect(isCustomConfigured({ apiKey: '', model: 'm', baseUrl: 'http://x/v1' })).toBe(true)
    expect(isCustomConfigured({ apiKey: 'k', model: '', baseUrl: 'http://x/v1' })).toBe(false)
    expect(isCustomConfigured({ apiKey: 'k', model: 'm', baseUrl: '' })).toBe(false)
    expect(isCustomConfigured({ apiKey: 'k', model: 'm', baseUrl: '   ' })).toBe(false)
    expect(isCustomConfigured(undefined)).toBe(false)
  })

  it('does not require an API key: local model servers accept anonymous requests', () => {
    expect(
      isCustomConfigured({ apiKey: '', model: 'llama3', baseUrl: 'http://localhost:11434/v1' }),
    ).toBe(true)
  })
})

describe('activeProvider', () => {
  it('uses the custom endpoint when it is selected and complete', () => {
    expect(activeProvider(withCustom({}))).toBe('custom')
  })

  it('falls back to genspark when the selected custom endpoint is incomplete', () => {
    expect(activeProvider(withCustom({ baseUrl: '' }))).toBe('genspark')
    expect(activeProvider(withCustom({ model: '' }))).toBe('genspark')
  })

  it('ignores a configured custom endpoint that is not selected', () => {
    const settings = withCustom({})
    settings.provider = 'genspark'
    expect(activeProvider(settings)).toBe('genspark')
  })

  it('normalizes the other vendor providers back to genspark', () => {
    const settings = defaultAiSettings()
    settings.provider = 'anthropic'
    settings.providers.anthropic.apiKey = 'sk-ant-whatever'
    expect(activeProvider(settings)).toBe('genspark')
  })
})

describe('toModelSettings', () => {
  it('flattens a live custom endpoint', () => {
    expect(toModelSettings(withCustom({ apiKey: 'sk-1' }))).toEqual({
      mode: 'custom',
      baseUrl: 'http://localhost:11434/v1',
      model: 'local-model',
      apiKey: 'sk-1',
      temperature: null,
      maxTokens: null,
      reasoningEffort: null,
    })
  })

  it('reports genspark mode but still returns the stored draft endpoint', () => {
    const settings = withCustom({ model: '' })
    expect(toModelSettings(settings)).toEqual({
      mode: 'genspark',
      baseUrl: 'http://localhost:11434/v1',
      model: '',
      apiKey: '',
      temperature: null,
      maxTokens: null,
      reasoningEffort: null,
    })
  })
})

describe('applyModelSettings', () => {
  it('stores a custom endpoint and selects it', () => {
    const next = applyModelSettings(defaultAiSettings(), {
      mode: 'custom',
      baseUrl: ' https://api.deepseek.com/v1 ',
      model: ' deepseek-chat ',
      apiKey: ' sk-2 ',
      temperature: null,
      maxTokens: null,
      reasoningEffort: null,
    })
    expect(next.provider).toBe('custom')
    // surrounding whitespace from a paste never reaches the request; the unset
    // knobs are omitted entirely rather than stored as nulls, except
    // temperature where null is itself the instruction "don't send it"
    expect(next.providers.custom).toEqual({
      baseUrl: 'https://api.deepseek.com/v1',
      model: 'deepseek-chat',
      apiKey: 'sk-2',
      temperature: null,
    })
  })

  it('keeps the typed endpoint when switching back to genspark', () => {
    const next = applyModelSettings(defaultAiSettings(), {
      mode: 'genspark',
      baseUrl: 'https://api.deepseek.com/v1',
      model: 'deepseek-chat',
      apiKey: 'sk-2',
      temperature: null,
      maxTokens: null,
      reasoningEffort: null,
    })
    expect(next.provider).toBe('genspark')
    expect(next.providers.custom.model).toBe('deepseek-chat')
  })

  it('refuses to select an incomplete custom endpoint, so AI keeps working', () => {
    const next = applyModelSettings(defaultAiSettings(), {
      mode: 'custom',
      baseUrl: 'https://api.deepseek.com/v1',
      model: '',
      apiKey: '',
      temperature: null,
      maxTokens: null,
      reasoningEffort: null,
    })
    expect(next.provider).toBe('genspark')
  })

  it('leaves the other providers untouched', () => {
    const settings = defaultAiSettings({ anthropic: 'preset-key' })
    const next = applyModelSettings(settings, {
      mode: 'custom',
      baseUrl: 'http://localhost:1234/v1',
      model: 'm',
      apiKey: '',
      temperature: null,
      maxTokens: null,
      reasoningEffort: null,
    })
    expect(next.providers.anthropic).toEqual(settings.providers.anthropic)
  })

  it('round-trips through toModelSettings', () => {
    const input: AiModelSettings = {
      mode: 'custom',
      baseUrl: 'http://localhost:8000/v1',
      model: 'qwen2.5',
      apiKey: '',
      temperature: null,
      maxTokens: null,
      reasoningEffort: null,
    }
    expect(toModelSettings(applyModelSettings(defaultAiSettings(), input))).toEqual(input)
  })

  it('round-trips the generation knobs too', () => {
    const input: AiModelSettings = {
      mode: 'custom',
      baseUrl: 'http://localhost:8000/v1',
      model: 'qwen2.5',
      apiKey: '',
      temperature: 0.9,
      maxTokens: 32_000,
      reasoningEffort: 'high',
    }
    const next = applyModelSettings(defaultAiSettings(), input)
    expect(next.providers.custom).toMatchObject({
      temperature: 0.9,
      maxTokens: 32_000,
      reasoningEffort: 'high',
    })
    expect(toModelSettings(next)).toEqual(input)
  })
})

describe('generation knobs', () => {
  const base: AiProviderConfig = { apiKey: '', model: 'm', baseUrl: 'http://x/v1' }

  it('sends the house default temperature when nothing is configured', () => {
    expect(temperatureField(base)).toEqual({ temperature: DEFAULT_TEMPERATURE })
  })

  it('omits temperature entirely when it is null — the reasoning-model case', () => {
    expect(temperatureField({ ...base, temperature: null })).toEqual({})
  })

  it('sends an explicit temperature, including 0 and 1', () => {
    expect(temperatureField({ ...base, temperature: 0 })).toEqual({ temperature: 0 })
    expect(temperatureField({ ...base, temperature: 1 })).toEqual({ temperature: 1 })
  })

  it('prefers the configured token ceiling over the caller default', () => {
    expect(resolveMaxTokens(base, 8192)).toBe(8192)
    expect(resolveMaxTokens({ ...base, maxTokens: 2048 }, 8192)).toBe(2048)
    // a nonsense stored value must not silently produce a zero-token request
    expect(resolveMaxTokens({ ...base, maxTokens: 0 }, 8192)).toBe(8192)
  })

  it('sends reasoning_effort only when set', () => {
    expect(reasoningEffortField(base)).toEqual({})
    expect(reasoningEffortField({ ...base, reasoningEffort: 'low' })).toEqual({
      reasoning_effort: 'low',
    })
  })

  it('recognizes exactly the four documented effort levels', () => {
    for (const effort of ['minimal', 'low', 'medium', 'high']) {
      expect(isReasoningEffort(effort)).toBe(true)
    }
    expect(isReasoningEffort('extreme')).toBe(false)
    expect(isReasoningEffort(null)).toBe(false)
  })
})

describe('maxTokensField', () => {
  const base: AiProviderConfig = { apiKey: '', model: 'm', baseUrl: 'http://x/v1' }

  it('stays silent when no ceiling is configured, leaving it to the server', () => {
    expect(maxTokensField(base)).toEqual({})
    expect(maxTokensField({ ...base, maxTokens: 0 })).toEqual({})
  })

  it('sends the configured ceiling', () => {
    expect(maxTokensField({ ...base, maxTokens: 4096 })).toEqual({ max_tokens: 4096 })
  })
})
