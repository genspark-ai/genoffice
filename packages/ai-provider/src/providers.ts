import type {
  AiCustomProfile,
  AiModelSettings,
  AiProviderConfig,
  AiProviderId,
  AiProviderMeta,
  AiSettings,
  LegacyAiSettings,
} from './types'

/**
 * Genspark server-side LLM proxy endpoints. All three protocols share the
 * api_key from the gsk login; model ids follow the proxy's own naming scheme,
 * which differs from the official vendor ids.
 */
export const GENSPARK_LLM_BASE_URLS = {
  anthropic: 'https://www.genspark.ai/api/anthropic',
  gemini: 'https://www.genspark.ai/api/llm_proxy/gemini/v1beta',
  openai: 'https://www.genspark.ai/api/llm_proxy/v1',
} as const

/**
 * Splits GenOffice usage out of the proxy's default "Claw" billing bucket
 * (the backend attributes gsk-key traffic by X-Agent-Type). Only sent to the
 * Genspark proxy — never to direct vendor APIs.
 */
export const GENSPARK_AGENT_TYPE = 'genoffice'

export function gensparkAttributionHeaders(baseUrl?: string): Record<string, string> {
  return baseUrl?.startsWith('https://www.genspark.ai')
    ? { 'X-Agent-Type': GENSPARK_AGENT_TYPE }
    : {}
}

export const AI_PROVIDERS: AiProviderMeta[] = [
  {
    id: 'genspark',
    label: 'Genspark',
    models: [
      'claude-opus-4-7',
      'claude-opus-4-8',
      'claude-sonnet-4-6',
      'claude-haiku-4-5',
      'gpt-5.2',
      'gemini-3.1-pro-preview',
      'gemini-3-flash-preview',
    ],
    defaultModel: 'claude-opus-4-7',
    keyPlaceholder: 'Not required - sign in to Genspark',
  },
  {
    id: 'anthropic',
    label: 'Claude',
    models: [
      'claude-sonnet-5',
      'claude-opus-4-8',
      'claude-opus-4-7',
      'claude-sonnet-4-6',
      'claude-opus-4-6',
      'claude-opus-4-5-20251101',
      'claude-haiku-4-5-20251001',
      'claude-sonnet-4-5-20250929',
    ],
    defaultModel: 'claude-opus-4-7',
    keyPlaceholder: 'sk-ant-api03-...',
  },
  {
    id: 'gemini',
    label: 'Gemini',
    models: ['gemini-2.5-pro', 'gemini-2.5-flash', 'gemini-2.0-flash'],
    defaultModel: 'gemini-2.5-flash',
    keyPlaceholder: 'AIza...',
  },
  {
    id: 'deepseek',
    label: 'DeepSeek',
    models: ['deepseek-chat', 'deepseek-reasoner'],
    defaultModel: 'deepseek-chat',
    keyPlaceholder: 'sk-...',
  },
  {
    id: 'openai',
    label: 'OpenAI',
    models: ['gpt-4.1', 'gpt-4.1-mini', 'gpt-4o', 'gpt-4o-mini'],
    defaultModel: 'gpt-4.1-mini',
    keyPlaceholder: 'sk-...',
  },
  {
    id: 'custom',
    label: 'Custom',
    models: [],
    defaultModel: '',
    keyPlaceholder: 'API Key',
    needsBaseUrl: true,
  },
]

/**
 * Fresh settings with every provider's default model and an empty key,
 * except providers listed in `defaultApiKeys` (e.g. an app-specific
 * preconfigured Anthropic key). Callers own that policy; this package
 * has no hardcoded keys.
 */
export function defaultAiSettings(
  defaultApiKeys?: Partial<Record<AiProviderId, string>>,
): AiSettings {
  const providers = {} as AiSettings['providers']
  for (const meta of AI_PROVIDERS) {
    providers[meta.id] = {
      apiKey: defaultApiKeys?.[meta.id] ?? '',
      model: meta.defaultModel,
      baseUrl: meta.needsBaseUrl ? '' : undefined,
    }
  }
  return { provider: 'genspark', providers, tavilyApiKey: '', proxyUrl: '' }
}

/** proxy schemes undici's ProxyAgent and Chromium's proxyRules both understand */
const PROXY_SCHEMES = ['http:', 'https:', 'socks5:', 'socks4:']

/**
 * Validate and normalize a user-typed proxy URL. A bare `host:port` is read as
 * `http://host:port`, the common shape people paste from a proxy client.
 * Returns '' for anything unusable, so a typo degrades to "no proxy" rather
 * than breaking every outbound request with an unparseable dispatcher.
 */
export function normalizeProxyUrl(raw: string | undefined): string {
  const value = (raw ?? '').trim()
  if (!value) return ''
  const withScheme = /^[a-z][a-z0-9+.-]*:\/\//i.test(value) ? value : `http://${value}`
  try {
    const url = new URL(withScheme)
    if (!PROXY_SCHEMES.includes(url.protocol.toLowerCase())) return ''
    if (!url.hostname) return ''
    return url.toString().replace(/\/$/, '')
  } catch {
    return ''
  }
}

/**
 * A custom endpoint is usable once it has somewhere to send the request and a
 * model to name in it. The API key stays optional: local servers (Ollama,
 * LM Studio, vLLM) accept anonymous requests.
 */
export function isCustomConfigured(config: AiProviderConfig | undefined): boolean {
  return !!config?.baseUrl?.trim() && !!config.model.trim()
}

/**
 * The provider a request actually goes to. Genspark (gsk login) is the default
 * and the fallback; a user-configured custom endpoint wins when it is both
 * selected and complete, so a half-filled settings form never silently breaks
 * AI features. Shared by every app's main process so docs, sheets, slides and
 * pdf always agree on which backend is live.
 */
export function activeProvider(settings: AiSettings): AiProviderId {
  return settings.provider === 'custom' && isCustomConfigured(settings.providers?.custom)
    ? 'custom'
    : 'genspark'
}

/** Strip the library-only fields: what reaches a request is a plain provider config. */
function profileConfig(profile: AiCustomProfile): AiProviderConfig {
  const { id: _id, label: _label, ...config } = profile
  return config
}

/** The profile `activeProfileId` names, falling back to the first saved one. */
export function activeProfile(settings: AiSettings): AiCustomProfile | undefined {
  const profiles = settings.customProfiles
  if (!profiles?.length) return undefined
  return profiles.find((p) => p.id === settings.activeProfileId) ?? profiles[0]
}

/**
 * Copy the selected profile into `providers.custom`, which is the only place
 * the request builders look. Keeping the indirection here means switching
 * models needs no change in `chatForProvider` / `streamForProvider` / any main
 * process, and a settings file written by an older build still resolves.
 */
export function syncActiveProfile(settings: AiSettings): AiSettings {
  const profile = activeProfile(settings)
  if (!profile) return settings
  return {
    ...settings,
    activeProfileId: profile.id,
    providers: { ...settings.providers, custom: profileConfig(profile) },
  }
}

/** Stable-enough id for a new profile; the settings file is single-writer. */
export function newProfileId(): string {
  return `p${Date.now().toString(36)}${Math.floor(Math.random() * 1e6).toString(36)}`
}

/** Flatten full settings into the shape the settings dialog edits. */
export function toModelSettings(settings: AiSettings): AiModelSettings {
  const live = activeProfile(settings)
  // the fields below describe the selected profile, so read them from it
  // rather than from providers.custom, which only mirrors it
  const custom = live ?? settings.providers?.custom
  return {
    mode: activeProvider(settings) === 'custom' ? 'custom' : 'genspark',
    profiles: (settings.customProfiles ?? []).map((p) => ({
      id: p.id,
      label: p.label,
      baseUrl: p.baseUrl ?? '',
      model: p.model,
      apiKey: p.apiKey,
    })),
    profileId: live?.id ?? null,
    baseUrl: custom?.baseUrl ?? '',
    model: custom?.model ?? '',
    apiKey: custom?.apiKey ?? '',
    // the dialog has no "unspecified" state: an absent stored value reads back
    // as null, i.e. "leave it to the model"
    temperature: custom?.temperature ?? null,
    maxTokens: custom?.maxTokens ?? null,
    reasoningEffort: custom?.reasoningEffort ?? null,
    // network settings are independent of the provider choice: Tavily and the
    // proxy apply whether calls go to Genspark or a custom endpoint
    tavilyApiKey: settings.tavilyApiKey ?? '',
    proxyUrl: settings.proxyUrl ?? '',
  }
}

/**
 * Fold the dialog's edits back into full settings. The custom endpoint is kept
 * even when the user switches back to Genspark, so toggling between the two
 * does not lose a typed-in configuration.
 */
export function applyModelSettings(settings: AiSettings, input: AiModelSettings): AiSettings {
  const custom: AiProviderConfig = {
    baseUrl: input.baseUrl.trim(),
    model: input.model.trim(),
    apiKey: input.apiKey.trim(),
    // null is meaningful here (omit the field), so it is stored as-is
    temperature: input.temperature,
    ...(input.maxTokens === null ? {} : { maxTokens: input.maxTokens }),
    ...(input.reasoningEffort === null ? {} : { reasoningEffort: input.reasoningEffort }),
  }
  // `input.profiles` is the library as the dialog now has it — rows the user
  // added, removed or renamed — and `input.profileId` says which one the
  // endpoint fields above describe. Existing rows keep their stored endpoint;
  // only the selected one takes the edits, so saving never disturbs the others.
  const stored = new Map((settings.customProfiles ?? []).map((p) => [p.id, p]))
  const profiles: AiCustomProfile[] = input.profiles.map((row) => {
    // the selected row takes the tuning knobs too; every other row keeps its
    // stored tuning and the endpoint the dialog is carrying for it
    if (row.id === input.profileId) return { ...custom, id: row.id, label: row.label.trim() }
    const previous = stored.get(row.id)
    return {
      ...previous,
      baseUrl: row.baseUrl.trim(),
      model: row.model.trim(),
      apiKey: row.apiKey.trim(),
      id: row.id,
      label: row.label.trim(),
    }
  })

  // A first-ever endpoint typed in without a row of its own still needs
  // somewhere to live, or it would vanish on the next read.
  if (!profiles.length && isCustomConfigured(custom)) {
    profiles.push({ ...custom, id: newProfileId(), label: custom.model })
  }
  const selected =
    profiles.find((p) => p.id === input.profileId) ?? (profiles.length ? profiles[0] : undefined)

  const next: AiSettings = {
    provider: input.mode === 'custom' ? 'custom' : 'genspark',
    providers: { ...settings.providers, custom },
    ...(profiles.length ? { customProfiles: profiles } : {}),
    ...(selected ? { activeProfileId: selected.id } : {}),
    tavilyApiKey: input.tavilyApiKey.trim(),
    proxyUrl: normalizeProxyUrl(input.proxyUrl),
  }
  // the live endpoint must follow the selection, not the last thing typed
  const settled = syncActiveProfile(next)
  // an incomplete custom endpoint would silently disable AI; keep Genspark live instead
  settled.provider = activeProvider(settled)
  return settled
}

/**
 * Merge on-disk settings over freshly computed defaults, migrating the
 * pre-provider shape (a single OpenAI-compatible endpoint) into the
 * "custom" provider slot. `stored` is whatever the caller read from its
 * settings file (already JSON-parsed); this function does no file I/O.
 */
export function resolveAiSettings(
  stored: Partial<AiSettings> & LegacyAiSettings,
  defaults: AiSettings,
): AiSettings {
  // network settings live outside the provider matrix, so they survive both
  // the legacy migration and the normal merge
  const network = {
    tavilyApiKey: stored.tavilyApiKey ?? defaults.tavilyApiKey ?? '',
    proxyUrl: normalizeProxyUrl(stored.proxyUrl ?? defaults.proxyUrl),
  }
  if (!stored.providers) {
    if (stored.apiKey) {
      defaults.providers.custom = {
        apiKey: stored.apiKey,
        model: stored.model ?? '',
        baseUrl: stored.baseUrl ?? 'https://api.openai.com/v1',
      }
    }
    return withProfiles({ ...defaults, ...network })
  }
  return withProfiles({
    provider: stored.provider ?? defaults.provider,
    providers: { ...defaults.providers, ...stored.providers },
    ...(stored.customProfiles ? { customProfiles: stored.customProfiles } : {}),
    ...(stored.activeProfileId ? { activeProfileId: stored.activeProfileId } : {}),
    ...network,
  })
}

/**
 * Settle the profile library. A file written before profiles existed has a
 * configured `providers.custom` and no list, so that endpoint becomes profile
 * #1 and stays live — an upgrade changes nothing the user can see. Once a list
 * exists the selected profile wins, since that is what the sidebar switches.
 */
function withProfiles(settings: AiSettings): AiSettings {
  if (settings.customProfiles?.length) return syncActiveProfile(settings)
  const custom = settings.providers?.custom
  if (!isCustomConfigured(custom)) return settings
  const profile: AiCustomProfile = {
    ...(custom as AiProviderConfig),
    id: newProfileId(),
    label: custom?.model?.trim() ?? '',
  }
  return { ...settings, customProfiles: [profile], activeProfileId: profile.id }
}
