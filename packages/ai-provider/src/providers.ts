import type {
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
  return { provider: 'genspark', providers }
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

/** Flatten full settings into the shape the settings dialog edits. */
export function toModelSettings(settings: AiSettings): AiModelSettings {
  const custom = settings.providers?.custom
  return {
    mode: activeProvider(settings) === 'custom' ? 'custom' : 'genspark',
    baseUrl: custom?.baseUrl ?? '',
    model: custom?.model ?? '',
    apiKey: custom?.apiKey ?? '',
    // the dialog has no "unspecified" state: an absent stored value reads back
    // as null, i.e. "leave it to the model"
    temperature: custom?.temperature ?? null,
    maxTokens: custom?.maxTokens ?? null,
    reasoningEffort: custom?.reasoningEffort ?? null,
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
  const next: AiSettings = {
    provider: input.mode === 'custom' ? 'custom' : 'genspark',
    providers: { ...settings.providers, custom },
  }
  // an incomplete custom endpoint would silently disable AI; keep Genspark live instead
  next.provider = activeProvider(next)
  return next
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
  if (!stored.providers) {
    if (stored.apiKey) {
      defaults.providers.custom = {
        apiKey: stored.apiKey,
        model: stored.model ?? '',
        baseUrl: stored.baseUrl ?? 'https://api.openai.com/v1',
      }
    }
    return defaults
  }
  return {
    provider: stored.provider ?? defaults.provider,
    providers: { ...defaults.providers, ...stored.providers },
  }
}
