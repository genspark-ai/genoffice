import type { AiProviderId, AiProviderMeta, AiSettings, LegacyAiSettings } from './types'

export const OPENROUTER_BASE_URL = 'https://openrouter.ai/api/v1'

/** Legacy Genspark endpoints retained only for saved-settings compatibility. */
export const GENSPARK_LLM_BASE_URLS = {
  anthropic: 'https://www.genspark.ai/api/anthropic',
  gemini: 'https://www.genspark.ai/api/llm_proxy/gemini/v1beta',
  openai: 'https://www.genspark.ai/api/llm_proxy/v1',
} as const
export const GENSPARK_AGENT_TYPE = 'genoffice'
export function gensparkAttributionHeaders(baseUrl?: string): Record<string, string> {
  return baseUrl?.startsWith('https://www.genspark.ai') ? { 'X-Agent-Type': GENSPARK_AGENT_TYPE } : {}
}

export const AI_PROVIDERS: AiProviderMeta[] = [
  {
    id: 'openrouter', label: 'OpenRouter (NVIDIA Nemotron)',
    models: ['nvidia/nemotron-3-ultra-550b-a55b:free', 'nvidia/nemotron-3-super-120b-a12b:free', 'nvidia/nemotron-3-nano-30b-a3b:free', 'openrouter/free'],
    defaultModel: 'nvidia/nemotron-3-ultra-550b-a55b:free', keyPlaceholder: 'sk-or-v1-...',
  },
  {
    id: 'genspark', label: 'Genspark (Legacy)',
    models: ['claude-opus-4-7', 'claude-opus-4-8', 'claude-sonnet-4-6', 'claude-haiku-4-5', 'gpt-5.2', 'gemini-3.1-pro-preview', 'gemini-3-flash-preview'],
    defaultModel: 'claude-opus-4-7', keyPlaceholder: 'Not required - sign in to Genspark',
  },
  { id: 'anthropic', label: 'Claude', models: ['claude-sonnet-5', 'claude-opus-4-8', 'claude-opus-4-7', 'claude-sonnet-4-6'], defaultModel: 'claude-opus-4-7', keyPlaceholder: 'sk-ant-api03-...' },
  { id: 'gemini', label: 'Gemini', models: ['gemini-2.5-pro', 'gemini-2.5-flash', 'gemini-2.0-flash'], defaultModel: 'gemini-2.5-flash', keyPlaceholder: 'AIza...' },
  { id: 'deepseek', label: 'DeepSeek', models: ['deepseek-chat', 'deepseek-reasoner'], defaultModel: 'deepseek-chat', keyPlaceholder: 'sk-...' },
  { id: 'openai', label: 'OpenAI', models: ['gpt-4.1', 'gpt-4.1-mini', 'gpt-4o', 'gpt-4o-mini'], defaultModel: 'gpt-4.1-mini', keyPlaceholder: 'sk-...' },
  { id: 'custom', label: 'Custom', models: [], defaultModel: '', keyPlaceholder: 'API Key', needsBaseUrl: true },
]

export function defaultAiSettings(defaultApiKeys?: Partial<Record<AiProviderId, string>>): AiSettings {
  const providers = {} as AiSettings['providers']
  for (const meta of AI_PROVIDERS) providers[meta.id] = { apiKey: defaultApiKeys?.[meta.id] ?? '', model: meta.defaultModel, baseUrl: meta.needsBaseUrl ? '' : undefined }
  return { provider: 'openrouter', providers }
}

export function resolveAiSettings(stored: Partial<AiSettings> & LegacyAiSettings, defaults: AiSettings): AiSettings {
  if (!stored.providers) {
    if (stored.apiKey) defaults.providers.custom = { apiKey: stored.apiKey, model: stored.model ?? '', baseUrl: stored.baseUrl ?? 'https://api.openai.com/v1' }
    return defaults
  }
  return { provider: stored.provider ?? defaults.provider, providers: { ...defaults.providers, ...stored.providers } }
}
