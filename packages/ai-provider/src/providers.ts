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
      'gpt-5.6',
      'gpt-5.6-terra',
      'gpt-5.6-luna',
      'gemini-3.1-pro-preview',
      'gemini-3-flash-preview',
      'gemini-3.7-flash',
    ],
    defaultModel: 'claude-opus-4-7',
    keyPlaceholder: 'Not required - sign in to Genspark',
  },
  {
    id: 'anthropic',
    label: 'Claude',
    // current-generation ids per platform.claude.com models overview (2026-08)
    models: [
      'claude-opus-5',
      'claude-sonnet-5',
      'claude-fable-5',
      'claude-opus-4-8',
      'claude-opus-4-7',
      'claude-sonnet-4-6',
      'claude-haiku-4-5-20251001',
    ],
    defaultModel: 'claude-sonnet-5',
    keyPlaceholder: 'sk-ant-api03-...',
  },
  {
    id: 'gemini',
    label: 'Gemini',
    // 3.x lineup per ai.google.dev pricing/model pages (2026-08)
    models: [
      'gemini-3.1-pro-preview',
      'gemini-3.6-flash',
      'gemini-3.5-flash',
      'gemini-3.5-flash-lite',
    ],
    defaultModel: 'gemini-3.6-flash',
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
    // GPT-5.6 naming: the bare `gpt-5.6` alias routes to gpt-5.6-sol (flagship);
    // terra balances cost/intelligence, luna is the high-volume tier (2026-08)
    models: ['gpt-5.6', 'gpt-5.6-terra', 'gpt-5.6-luna', 'gpt-5.5', 'gpt-5.4', 'gpt-5.4-mini'],
    defaultModel: 'gpt-5.6-terra',
    keyPlaceholder: 'sk-...',
  },
  {
    id: 'kimi',
    label: 'Kimi',
    models: ['kimi-k3'],
    defaultModel: 'kimi-k3',
    keyPlaceholder: 'sk-...',
  },
  {
    id: 'glm',
    label: 'GLM',
    models: ['glm-5.3', 'glm-5-turbo', 'glm-4.7'],
    defaultModel: 'glm-5.3',
    keyPlaceholder: 'xxxxxxxx.xxxxxxxx',
  },
  {
    id: 'qwen',
    // DashScope "latest" aliases, stable across Qwen releases
    label: 'Qwen',
    models: ['qwen-max', 'qwen-plus', 'qwen-flash'],
    defaultModel: 'qwen-max',
    keyPlaceholder: 'sk-...',
  },
  {
    id: 'doubao',
    // Ark also accepts ep-... inference endpoint ids in the model field
    label: 'Doubao',
    models: ['doubao-seed-1-6-251015'],
    defaultModel: 'doubao-seed-1-6-251015',
    keyPlaceholder: 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx',
  },
  {
    id: 'minimax',
    label: 'MiniMax',
    models: ['MiniMax-M2.5'],
    defaultModel: 'MiniMax-M2.5',
    keyPlaceholder: 'eyJ...',
  },
  {
    id: 'xai',
    label: 'Grok',
    models: ['grok-4.6', 'grok-4.5'],
    defaultModel: 'grok-4.6',
    keyPlaceholder: 'xai-...',
  },
  {
    id: 'mistral',
    label: 'Mistral',
    models: ['mistral-large-latest', 'mistral-small-latest', 'codestral-latest'],
    defaultModel: 'mistral-large-latest',
    keyPlaceholder: 'API Key',
  },
  {
    id: 'openrouter',
    label: 'OpenRouter',
    models: [
      'openrouter/auto',
      'anthropic/claude-sonnet-5',
      'openai/gpt-5.6',
      'moonshotai/kimi-k3',
    ],
    defaultModel: 'openrouter/auto',
    keyPlaceholder: 'sk-or-...',
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
  return { provider: 'genspark', providers, gskToolsEnabled: true }
}

/** false only on an explicit opt-out; absent (pre-toggle settings files) means on */
export function cloudToolsEnabled(settings: Pick<AiSettings, 'gskToolsEnabled'>): boolean {
  return settings.gskToolsEnabled !== false
}

/**
 * The stored provider selection is honored only when its config is usable
 * (api-key providers need a key and a model id; providers flagged
 * needsBaseUrl also need a base URL). Anything else — including unknown
 * ids from a hand-edited
 * settings file — falls back to genspark, so a half-filled setup degrades
 * to the signed-in default instead of silently disabling AI.
 */
export function activeProvider(settings: AiSettings): AiProviderId {
  const provider = settings.provider
  if (provider === 'genspark') return 'genspark'
  const meta = AI_PROVIDERS.find((m) => m.id === provider)
  const config = settings.providers?.[provider]
  if (!meta || !config?.apiKey || !config.model) return 'genspark'
  if (meta.needsBaseUrl && !config.baseUrl) return 'genspark'
  return provider
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
    gskToolsEnabled: stored.gskToolsEnabled ?? defaults.gskToolsEnabled ?? true,
  }
}
