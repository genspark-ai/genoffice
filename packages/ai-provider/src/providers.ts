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
    // 3.x lineup per ai.google.dev/gemini-api/docs/models (2026-08). 3.7 Flash is
    // the current stable Flash; 3.1 Pro is still preview-only.
    models: [
      'gemini-3.7-flash',
      'gemini-3.1-pro-preview',
      'gemini-3.6-flash',
      'gemini-3.5-flash-lite',
    ],
    defaultModel: 'gemini-3.7-flash',
    keyPlaceholder: 'AIza...',
  },
  {
    id: 'deepseek',
    label: 'DeepSeek',
    // V4 ids per api-docs.deepseek.com (2026-08). The deepseek-chat /
    // deepseek-reasoner aliases were retired 2026-07-24; thinking mode is now
    // a request parameter, so both ids drive the tool-calling agent loop.
    models: ['deepseek-v4-pro', 'deepseek-v4-flash'],
    defaultModel: 'deepseek-v4-pro',
    keyPlaceholder: 'sk-...',
  },
  {
    id: 'openai',
    label: 'OpenAI',
    // GPT-5.6 naming: sol is the flagship (the bare `gpt-5.6` alias resolves to
    // it, but spell it out so the picker says which tier it is), terra balances
    // cost/intelligence, luna is the high-volume tier (2026-08)
    models: ['gpt-5.6-sol', 'gpt-5.6-terra', 'gpt-5.6-luna', 'gpt-5.5', 'gpt-5.4', 'gpt-5.4-mini'],
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
    // bigmodel.cn text-model lineup (2026-08); 5.3 and 5.2 share a base model,
    // 5-Turbo is the cheap tier
    models: ['glm-5.3', 'glm-5.2', 'glm-5-turbo'],
    defaultModel: 'glm-5.3',
    keyPlaceholder: 'xxxxxxxx.xxxxxxxx',
  },
  {
    id: 'qwen',
    label: 'Qwen',
    // Versioned DashScope ids: the bare qwen-max alias still points at a
    // Qwen2.5-era snapshot, so name the 3.x tiers explicitly (2026-08)
    models: ['qwen3.8-max', 'qwen3.7-plus', 'qwen3.7-flash'],
    defaultModel: 'qwen3.8-max',
    keyPlaceholder: 'sk-...',
  },
  {
    id: 'doubao',
    label: 'Doubao',
    // Ark ids are dashed and date-pinned; it also accepts ep-... inference
    // endpoint ids in the model field
    models: ['doubao-seed-2-1-pro-260628', 'doubao-seed-2-1-turbo-260628'],
    defaultModel: 'doubao-seed-2-1-pro-260628',
    keyPlaceholder: 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx',
  },
  {
    id: 'minimax',
    label: 'MiniMax',
    // M3 is the current agentic/tool-use model; M2.5 moved to the legacy tier
    models: ['MiniMax-M3', 'MiniMax-M2.7'],
    defaultModel: 'MiniMax-M3',
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
    // `-latest` aliases track the newest GA snapshot. Medium 3.5 is Mistral's
    // agentic tier; codestral is a code-completion/FIM model, not an agent driver.
    models: ['mistral-medium-latest', 'mistral-large-latest', 'mistral-small-latest'],
    defaultModel: 'mistral-medium-latest',
    keyPlaceholder: 'API Key',
  },
  {
    id: 'openrouter',
    label: 'OpenRouter',
    // vendor-prefixed slugs exactly as openrouter.ai/api/v1/models lists them —
    // there is no `openai/gpt-5.6` alias there, only the per-tier ids
    models: [
      'openrouter/auto',
      'anthropic/claude-sonnet-5',
      'openai/gpt-5.6-sol',
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
 * Model ids a vendor has stopped serving, mapped to their replacement. A
 * stored selection outlives the provider list, so without this remap an old
 * settings file keeps sending an id the API now rejects.
 */
const RETIRED_MODELS: Partial<Record<AiProviderId, Record<string, string>>> = {
  // aliases retired 2026-07-24; DeepSeek pointed both at the V4-Flash line,
  // where thinking mode is a request parameter rather than a separate id
  deepseek: {
    'deepseek-chat': 'deepseek-v4-flash',
    'deepseek-reasoner': 'deepseek-v4-flash',
  },
}

/** pasted keys/URLs often carry stray whitespace, which turns into a 401 with a valid key */
function trimConfigs(providers: AiSettings['providers']): AiSettings['providers'] {
  const trimmed = { ...providers }
  for (const [id, config] of Object.entries(trimmed)) {
    trimmed[id as AiProviderId] = {
      ...config,
      apiKey: config.apiKey?.trim() ?? '',
      ...(config.baseUrl !== undefined ? { baseUrl: config.baseUrl.trim() } : {}),
    }
  }
  return trimmed
}

function migrateRetiredModels(providers: AiSettings['providers']): AiSettings['providers'] {
  const migrated = { ...providers }
  for (const [id, replacements] of Object.entries(RETIRED_MODELS)) {
    const config = migrated[id as AiProviderId]
    const replacement = config?.model ? replacements[config.model] : undefined
    if (replacement) migrated[id as AiProviderId] = { ...config, model: replacement }
  }
  return migrated
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
        apiKey: stored.apiKey.trim(),
        model: stored.model ?? '',
        baseUrl: (stored.baseUrl ?? 'https://api.openai.com/v1').trim(),
      }
    }
    return defaults
  }
  return {
    provider: stored.provider ?? defaults.provider,
    providers: trimConfigs(migrateRetiredModels({ ...defaults.providers, ...stored.providers })),
    gskToolsEnabled: stored.gskToolsEnabled ?? defaults.gskToolsEnabled ?? true,
  }
}
