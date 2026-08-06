import type { AiProviderId, AiProviderMeta, AiSettings, LegacyAiSettings } from './types'

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
    models: [
      'gemini-2.5-flash',
      'gemini-2.5-flash-lite',
      'gemini-3-flash-preview',
      'gemini-3.1-flash-lite',
      'gemini-3.5-flash',
      'gemini-2.5-pro',
      'gemini-3.1-pro-preview',
    ],
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
    id: 'openrouter',
    label: 'OpenRouter',
    models: [
      'openrouter/free',
      'anthropic/claude-sonnet-4.5',
      'openai/gpt-5.1',
      'google/gemini-3.1-pro-preview',
      'deepseek/deepseek-v3.2',
      'meta-llama/llama-3.3-70b-instruct',
      'qwen/qwen-2.5-72b-instruct',
      'mistralai/mistral-small-3.1',
    ],
    defaultModel: 'openrouter/free',
    keyPlaceholder: 'sk-or-v1-...',
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

/**
 * Bring-your-own-key / local-LLM overrides, applied on top of whatever provider
 * settings were resolved from disk. Reading these once from the environment lets
 * every app in the suite (docs, sheets, slides, pdf, shell) point at the same
 * endpoint with a single shared configuration — no per-app settings-file edits.
 *
 * Environment:
 *   GENOFFICE_AI_PROVIDER  provider id (default: 'custom' when a base URL is given)
 *   GENOFFICE_AI_BASE_URL  OpenAI-compatible base URL, e.g. http://localhost:11434/v1
 *   GENOFFICE_AI_MODEL     model name the endpoint exposes
 *   GENOFFICE_AI_API_KEY   api key (local servers usually accept any non-empty value) *
 * If GENOFFICE_AI_PROVIDER is omitted but GENOFFICE_AI_BASE_URL is set, the
 * active provider is switched to `custom` so a local/OpenAI-compatible endpoint
 * works out of the box. Genspark remains the default whenever nothing is set.
 */
export function applyProviderOverrides(settings: AiSettings): AiSettings {
  const envProvider = process.env.GENOFFICE_AI_PROVIDER as AiProviderId | undefined
  const baseUrl = process.env.GENOFFICE_AI_BASE_URL
  const model = process.env.GENOFFICE_AI_MODEL
  const apiKey = process.env.GENOFFICE_AI_API_KEY

  let provider = settings.provider
  if (envProvider) provider = envProvider
  else if (baseUrl && provider === 'genspark') provider = 'custom'

  const cfg = settings.providers[provider]
  if (!cfg) return settings

  settings.provider = provider
  if (model) cfg.model = model
  if (apiKey !== undefined) cfg.apiKey = apiKey
  if (provider === 'custom' && baseUrl) cfg.baseUrl = baseUrl
  return settings
}
