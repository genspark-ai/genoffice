import type { AgentMessage, AgentToolCall, AgentToolDef } from '@genoffice/agent-core'

/** Built-in provider presets. `custom` accepts any OpenAI-compatible endpoint. */
export type AiProviderId =
  | 'genspark'
  | 'anthropic'
  | 'gemini'
  | 'deepseek'
  | 'openai'
  | 'openrouter'
  | 'mistral'
  | 'groq'
  | 'together'
  | 'fireworks'
  | 'cerebras'
  | 'xai'
  | 'nvidia'
  | 'ollama'
  | 'lmstudio'
  | 'vllm'
  | 'llamacpp'
  | 'runware'
  | 'replicate'
  | 'fal'
  | 'stability'
  | 'custom'

export type AiProviderProtocol = 'genspark' | 'anthropic' | 'gemini' | 'openai-compatible'

export type AiEndpointKind = 'cloud' | 'local' | 'custom'

export type AiCapability = 'chat' | 'tools' | 'vision' | 'image-generation' | 'image-editing'

/** Genspark account status (gsk login state; the sole auth source for AI features) */
export interface GenSparkAccountStatus {
  loggedIn: boolean
  email?: string
}

export interface AiProviderConfig {
  apiKey: string
  model: string
  /** Used by custom and user-overridden OpenAI-compatible/local endpoints. */
  baseUrl?: string | undefined
  /** Optional provider-specific headers (never populated by discovery). */
  headers?: Record<string, string> | undefined
}

export interface AiProviderMeta {
  id: AiProviderId
  label: string
  models: string[]
  defaultModel: string
  keyPlaceholder: string
  needsBaseUrl?: boolean
  protocol?: AiProviderProtocol
  defaultBaseUrl?: string
  endpointKind?: AiEndpointKind
  capabilities?: AiCapability[]
  /** Curated image model identifiers shown before a live catalog is available. */
  imageModels?: string[]
  /** Image requests use the provider's native endpoint or OpenAI image API. */
  imageProtocol?:
    'gemini' | 'openai-images' | 'runware' | 'replicate' | 'fal' | 'stability' | 'none'
  modelListPath?: string
}

export interface AiSettings {
  provider: AiProviderId
  providers: Record<AiProviderId, AiProviderConfig>
}

/** pre-provider settings shape (single OpenAI-compatible endpoint); migrated into "custom" */
export interface LegacyAiSettings {
  baseUrl?: string
  apiKey?: string
  model?: string
}

export interface AiChatRequest {
  settings: AiSettings
  system: string
  user: string
}

export interface AiChatResponse {
  ok: boolean
  content?: string
  error?: string
}

export interface AiStreamRequest {
  requestId: string
  /** Deprecated renderer-provided settings. Main-process callers should omit this. */
  settings?: AiSettings
  /** Main-process task selector used by the secure resolver. */
  task?: 'chat' | 'image' | 'vision' | 'slides-generation' | (string & {})
  system: string
  messages: AgentMessage[]
  tools?: AgentToolDef[]
  maxTokens?: number
}

export interface AiStreamChunk {
  requestId: string
  /** 'ping' = wire-level keepalive so the renderer can tell a live stream from a dead one */
  type: 'delta' | 'tool-call' | 'done' | 'error' | 'ping'
  text?: string
  /** complete parsed tool call (emitted once its arguments finish streaming) */
  toolCall?: AgentToolCall
  error?: string
  /** machine-readable error cause ('timeout', exhausted 'credits'); lets the renderer localize the message */
  errorCode?: 'timeout' | 'credits'
  /** normalized stop reason carried on 'done' ('max_tokens' = output cut off by the token limit) */
  stopReason?: string
}

export interface AiModelCapabilities {
  chat?: boolean
  tools?: boolean
  vision?: boolean
  imageGeneration?: boolean
  imageEditing?: boolean
}

export interface AiModel {
  id: string
  displayName?: string
  createdAt?: number
  contextLength?: number
  inputModalities?: string[]
  outputModalities?: string[]
  capabilities: AiModelCapabilities
  providerMetadata?: Record<string, unknown>
}

export interface AiModelCatalog {
  provider: AiProviderId
  baseUrl?: string
  models: AiModel[]
  source: 'remote' | 'fallback'
  fetchedAt: number
}

export interface AiModelDiscoveryOptions {
  baseUrl?: string
  timeoutMs?: number
  signal?: AbortSignal
  /** Provider-side search term when the catalog supports it. */
  search?: string
  /** Limit pagination for quick connection/model probes. */
  maxPages?: number
}

export interface AiEndpointPolicy {
  /** Allow loopback, RFC1918, link-local and `.local` endpoints. */
  allowLocal: boolean
  /** Allow non-TLS HTTP endpoints. This should normally only be true for local hosts. */
  allowInsecureHttp: boolean
  /** Permit credentials or query strings in endpoint URLs. Disabled by default. */
  allowUrlCredentials: boolean
}

export interface AiEndpointValidation {
  ok: boolean
  normalized?: string
  kind?: AiEndpointKind
  reason?: string
}

export interface AiImageGenerationRequest {
  prompt: string
  model?: string
  count?: number
  size?: string
  aspectRatio?: string
  imageSize?: string
}

export interface AiGeneratedImage {
  provider: AiProviderId
  model: string
  mimeType: string
  base64?: string
  url?: string
  revisedPrompt?: string
}

export interface AiImageGenerationOptions {
  baseUrl?: string
  timeoutMs?: number
  signal?: AbortSignal
}
