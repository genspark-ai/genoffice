import { AI_PROVIDERS, type AiProviderId, type AiProviderMeta } from '@genoffice/ai-provider'

/**
 * Public shell contract for the AI Providers settings screen.
 *
 * Provider credentials are deliberately write-only: the renderer can submit a
 * new credential, but a read never returns it. The provider runtime owns the
 * actual credential store and can replace the shell IPC implementation without
 * changing this UI contract.
 */

export type AiProviderProtocol = 'genspark' | 'native' | 'openai-compatible' | 'local'

export interface AiProviderDefinition {
  id: string
  label: string
  description: string
  protocol: AiProviderProtocol
  /** Short identifier used by the renderer's local, non-network icon map. */
  icon: string
  requiresApiKey: boolean
  needsBaseUrl: boolean
  supportsModelDiscovery: boolean
  /** Defaults to true; image-only services set this to false. */
  supportsText?: boolean
  supportsImages: boolean
  defaultBaseUrl?: string
  /** Text/chat models. */
  models: string[]
  /** Image models shown on the Image tab. */
  imageModels?: string[]
}

export interface AiProviderConfigView {
  providerId: string
  model: string
  baseUrl: string
  credentialSet: boolean
  /** Renderer-safe masked suffix such as "••••a1b2"; never contains the full key. */
  credentialHint?: string
  enabled: boolean
}

export interface AiSettingsSnapshot {
  activeProvider: string
  activeModel: string
  imageProvider: string
  imageModel: string
  providers: AiProviderConfigView[]
  definitions: AiProviderDefinition[]
}

/** API-key fields are accepted only on writes and are never part of a snapshot. */
export interface SaveAiProviderInput {
  providerId: string
  capability?: AiProviderCapability
  /** `discover` refreshes the selector; `test` probes the chosen model. */
  operation?: 'discover' | 'test'
  model: string
  baseUrl?: string
  apiKey?: string
  clearCredential?: boolean
  enabled?: boolean
}

export interface SaveAiSettingsInput {
  activeProvider?: string
  activeModel?: string
  imageProvider?: string
  imageModel?: string
  provider?: SaveAiProviderInput
}

export interface AiProviderConnectionResult {
  ok: boolean
  message: string
  /** The runtime may return a discovered model list after a successful probe. */
  models?: string[]
}

export type AiProviderCapability = 'text' | 'image'

const PROVIDER_DESCRIPTIONS: Record<AiProviderId, string> = {
  genspark: 'Use the signed-in Genspark account.',
  openai: 'OpenAI models through the official API.',
  anthropic: 'Claude models through the Anthropic API.',
  gemini: 'Google Gemini models and image generation.',
  deepseek: 'DeepSeek chat and reasoning models.',
  openrouter: 'Route chat and image requests through OpenRouter.',
  mistral: 'Mistral models through an OpenAI-compatible API.',
  groq: 'Low-latency inference through Groq.',
  together: 'Open models through Together AI.',
  fireworks: 'Open-model inference and tool calling through Fireworks AI.',
  cerebras: 'Cerebras inference through an OpenAI-compatible API.',
  nvidia: 'NVIDIA-hosted or self-hosted NIM endpoints.',
  xai: 'Grok models through the xAI API.',
  ollama: 'Run local models through Ollama.',
  lmstudio: 'Use a local LM Studio server.',
  vllm: 'Use a local or hosted vLLM OpenAI-compatible server.',
  llamacpp: 'Use a local llama.cpp OpenAI-compatible server.',
  runware: 'Image generation through the Runware model catalog.',
  replicate: 'Official and community image models hosted by Replicate.',
  fal: 'Image models through the fal queued inference API.',
  stability: 'Stable Image Core, Ultra, and Stable Diffusion 3.5.',
  custom: 'Connect any OpenAI-compatible endpoint.',
}

const PROVIDER_ICONS: Partial<Record<AiProviderId, string>> = {
  lmstudio: 'lm-studio',
  llamacpp: 'llama',
}

function shellProtocol(provider: AiProviderMeta): AiProviderProtocol {
  if (provider.id === 'genspark') return 'genspark'
  if (provider.endpointKind === 'local') return 'local'
  if (provider.protocol === 'anthropic' || provider.protocol === 'gemini') return 'native'
  if (['runware', 'replicate', 'fal', 'stability'].includes(provider.id)) return 'native'
  return 'openai-compatible'
}

function requiresApiKey(provider: AiProviderMeta): boolean {
  return provider.id !== 'genspark' && provider.id !== 'custom' && provider.endpointKind !== 'local'
}

/**
 * The runtime registry owns endpoints and model presets. The shell adds only
 * presentation metadata so model lists cannot drift between the UI and runtime.
 */
export const AI_PROVIDER_DEFINITIONS: readonly AiProviderDefinition[] = AI_PROVIDERS.map(
  (provider) => {
    const supportsText = provider.capabilities?.includes('chat') ?? true
    const supportsImages = Boolean(provider.imageModels?.length)
    return {
      id: provider.id,
      label: provider.label,
      description: PROVIDER_DESCRIPTIONS[provider.id],
      protocol: shellProtocol(provider),
      icon: PROVIDER_ICONS[provider.id] ?? provider.id,
      requiresApiKey: requiresApiKey(provider),
      needsBaseUrl:
        provider.id === 'nvidia' || provider.endpointKind === 'local' || provider.id === 'custom',
      supportsModelDiscovery: provider.id !== 'genspark' && provider.id !== 'stability',
      supportsText,
      supportsImages,
      ...(provider.defaultBaseUrl ? { defaultBaseUrl: provider.defaultBaseUrl } : {}),
      models: supportsText ? [...provider.models] : [],
      ...(provider.imageModels ? { imageModels: [...provider.imageModels] } : {}),
    }
  },
)

export const AI_SETTINGS_CHANNELS = {
  get: 'ai-settings:get',
  save: 'ai-settings:save',
  test: 'ai-settings:test',
  cancelTest: 'ai-settings:cancel-test',
} as const

export interface AiSettingsApi {
  get(): Promise<AiSettingsSnapshot>
  save(input: SaveAiSettingsInput): Promise<AiSettingsSnapshot>
  test(input: SaveAiProviderInput): Promise<AiProviderConnectionResult>
  cancelTest(): Promise<void>
}
