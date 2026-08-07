import { describe, expect, it } from 'vitest'
import {
  AI_PROVIDERS,
  activeProfile,
  activeProvider,
  applyModelSettings,
  defaultAiSettings,
  isCustomConfigured,
  normalizeProxyUrl,
  resolveAiSettings,
  syncActiveProfile,
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
      tavilyApiKey: '',
      proxyUrl: '',
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
      tavilyApiKey: '',
      proxyUrl: '',
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
      tavilyApiKey: '',
      proxyUrl: '',
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
      tavilyApiKey: '',
      proxyUrl: '',
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
      tavilyApiKey: '',
      proxyUrl: '',
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
      tavilyApiKey: '',
      proxyUrl: '',
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
      tavilyApiKey: '',
      proxyUrl: '',
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
      tavilyApiKey: '',
      proxyUrl: '',
    }
    const next = applyModelSettings(defaultAiSettings(), input)
    // the knobs belong to the provider config; Tavily and the proxy are
    // top-level network settings, independent of which provider is selected
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

describe('normalizeProxyUrl', () => {
  it('accepts the schemes undici and Chromium both understand', () => {
    expect(normalizeProxyUrl('http://127.0.0.1:7897')).toBe('http://127.0.0.1:7897')
    expect(normalizeProxyUrl('https://proxy.example.com:8443')).toBe(
      'https://proxy.example.com:8443',
    )
    expect(normalizeProxyUrl('socks5://127.0.0.1:1080')).toBe('socks5://127.0.0.1:1080')
    expect(normalizeProxyUrl('socks4://127.0.0.1:1080')).toBe('socks4://127.0.0.1:1080')
  })

  it('reads a bare host:port as http, the shape people paste from a proxy client', () => {
    expect(normalizeProxyUrl('127.0.0.1:7897')).toBe('http://127.0.0.1:7897')
  })

  it('keeps credentials, which authenticated corporate proxies need', () => {
    expect(normalizeProxyUrl('http://user:pass@proxy:3128')).toBe('http://user:pass@proxy:3128')
  })

  it('degrades a typo to no-proxy rather than breaking every request', () => {
    expect(normalizeProxyUrl('ftp://nope:21')).toBe('')
    expect(normalizeProxyUrl('http://')).toBe('')
    expect(normalizeProxyUrl('   ')).toBe('')
    expect(normalizeProxyUrl(undefined)).toBe('')
  })
})

describe('network settings persistence', () => {
  it('round-trips the Tavily key and proxy through the dialog projection', () => {
    const next = applyModelSettings(defaultAiSettings(), {
      mode: 'genspark',
      baseUrl: '',
      model: '',
      apiKey: '',
      temperature: null,
      maxTokens: null,
      reasoningEffort: null,
      tavilyApiKey: '  tvly-abc  ',
      proxyUrl: '127.0.0.1:7897',
    })
    expect(next.tavilyApiKey).toBe('tvly-abc')
    expect(next.proxyUrl).toBe('http://127.0.0.1:7897')
    expect(toModelSettings(next)).toMatchObject({
      tavilyApiKey: 'tvly-abc',
      proxyUrl: 'http://127.0.0.1:7897',
    })
  })

  it('survives the legacy single-endpoint migration', () => {
    const resolved = resolveAiSettings(
      { apiKey: 'legacy', tavilyApiKey: 'tvly-x', proxyUrl: 'socks5://127.0.0.1:1080' },
      defaultAiSettings(),
    )
    expect(resolved.tavilyApiKey).toBe('tvly-x')
    expect(resolved.proxyUrl).toBe('socks5://127.0.0.1:1080')
  })
})

// ────────────────────────────────────────────────────────────
// Custom model profiles (sidebar switcher)
// ────────────────────────────────────────────────────────────

const ollama = {
  id: 'p-ollama',
  label: 'Local Ollama',
  baseUrl: 'http://localhost:11434/v1',
  model: 'qwen3',
  apiKey: '',
}
const gateway = {
  id: 'p-gateway',
  label: 'Work gateway',
  baseUrl: 'https://gw.example.com/v1',
  model: 'gpt-5.2',
  apiKey: 'sk-work',
  temperature: null,
}

describe('custom model profiles', () => {
  it('migrates a pre-profiles custom endpoint into profile #1 and keeps it live', () => {
    const settings = resolveAiSettings(
      {
        provider: 'custom',
        providers: {
          ...defaultAiSettings().providers,
          custom: {
            baseUrl: 'https://api.deepseek.com/v1',
            model: 'deepseek-chat',
            apiKey: 'sk-d',
          },
        },
      },
      defaultAiSettings(),
    )

    expect(settings.customProfiles).toHaveLength(1)
    expect(settings.customProfiles?.[0]).toMatchObject({
      baseUrl: 'https://api.deepseek.com/v1',
      model: 'deepseek-chat',
      // no name was ever typed, so the model id stands in
      label: 'deepseek-chat',
    })
    expect(settings.activeProfileId).toBe(settings.customProfiles?.[0]?.id)
    // the endpoint requests use is untouched by the migration
    expect(activeProvider(settings)).toBe('custom')
    expect(settings.providers.custom.model).toBe('deepseek-chat')
  })

  it('invents no profile when no custom endpoint was ever configured', () => {
    const settings = resolveAiSettings({}, defaultAiSettings())
    expect(settings.customProfiles).toBeUndefined()
    expect(activeProvider(settings)).toBe('genspark')
  })

  it('makes the selected profile the endpoint requests use', () => {
    const settings = resolveAiSettings(
      {
        provider: 'custom',
        providers: defaultAiSettings().providers,
        customProfiles: [ollama, gateway],
        activeProfileId: 'p-gateway',
      },
      defaultAiSettings(),
    )

    expect(settings.providers.custom).toEqual({
      baseUrl: 'https://gw.example.com/v1',
      model: 'gpt-5.2',
      apiKey: 'sk-work',
      temperature: null,
    })
    // id/label are library bookkeeping and must not reach a request
    expect(settings.providers.custom).not.toHaveProperty('id')
    expect(settings.providers.custom).not.toHaveProperty('label')
  })

  it('falls back to the first profile when the selected id is gone', () => {
    const settings = resolveAiSettings(
      {
        provider: 'custom',
        providers: defaultAiSettings().providers,
        customProfiles: [ollama, gateway],
        activeProfileId: 'p-deleted',
      },
      defaultAiSettings(),
    )

    expect(settings.activeProfileId).toBe('p-ollama')
    expect(settings.providers.custom.model).toBe('qwen3')
  })

  it('switching profiles swaps the live endpoint without touching the library', () => {
    const base = resolveAiSettings(
      {
        provider: 'custom',
        providers: defaultAiSettings().providers,
        customProfiles: [ollama, gateway],
        activeProfileId: 'p-ollama',
      },
      defaultAiSettings(),
    )

    const switched = syncActiveProfile({ ...base, activeProfileId: 'p-gateway' })

    expect(switched.providers.custom.model).toBe('gpt-5.2')
    expect(switched.customProfiles).toEqual(base.customProfiles)
    expect(activeProfile(switched)?.label).toBe('Work gateway')
  })

  it('leaves settings alone when the library is empty', () => {
    const settings = defaultAiSettings()
    expect(syncActiveProfile(settings)).toBe(settings)
    expect(activeProfile(settings)).toBeUndefined()
  })
})

describe('applyModelSettings with a profile library', () => {
  const stored = () =>
    resolveAiSettings(
      {
        provider: 'custom',
        providers: defaultAiSettings().providers,
        customProfiles: [ollama, gateway],
        activeProfileId: 'p-gateway',
      },
      defaultAiSettings(),
    )

  it('edits the selected profile instead of replacing the library', () => {
    const next = applyModelSettings(stored(), {
      ...toModelSettings(stored()),
      mode: 'custom',
      model: 'gpt-5.2-mini',
    })

    expect(next.customProfiles).toHaveLength(2)
    // the other model the user added survives a save
    expect(next.customProfiles?.find((p) => p.id === 'p-ollama')?.model).toBe('qwen3')
    const edited = next.customProfiles?.find((p) => p.id === 'p-gateway')
    expect(edited?.model).toBe('gpt-5.2-mini')
    // the name the user gave it is not clobbered by an endpoint edit
    expect(edited?.label).toBe('Work gateway')
    expect(next.providers.custom.model).toBe('gpt-5.2-mini')
  })

  it('keeps the library when the user switches back to Genspark', () => {
    const next = applyModelSettings(stored(), { ...toModelSettings(stored()), mode: 'genspark' })

    expect(next.provider).toBe('genspark')
    expect(next.customProfiles).toHaveLength(2)
  })

  it('seeds the library from a first-ever custom endpoint', () => {
    const next = applyModelSettings(defaultAiSettings(), {
      ...toModelSettings(defaultAiSettings()),
      mode: 'custom',
      baseUrl: 'http://localhost:1234/v1',
      model: 'llama4',
    })

    expect(next.customProfiles).toHaveLength(1)
    expect(next.customProfiles?.[0]).toMatchObject({ model: 'llama4', label: 'llama4' })
    expect(activeProvider(next)).toBe('custom')
  })
})
