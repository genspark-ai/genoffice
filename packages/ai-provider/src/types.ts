import type { AgentMessage, AgentToolCall, AgentToolDef } from '@genoffice/agent-core'

export type AiProviderId = 'genspark' | 'anthropic' | 'gemini' | 'deepseek' | 'openai' | 'custom'

/** Genspark account status (gsk login state; the sole auth source for AI features) */
export interface GenSparkAccountStatus {
  loggedIn: boolean
  email?: string
}

/** OpenAI-compatible `reasoning_effort` budget for models that think before answering */
export type ReasoningEffort = 'minimal' | 'low' | 'medium' | 'high'

export interface AiProviderConfig {
  apiKey: string
  model: string
  /** only used by the custom (OpenAI-compatible) provider */
  baseUrl?: string | undefined
  /**
   * Sampling temperature. `undefined` keeps the built-in default; `null` drops
   * the field from the request entirely — reasoning models (GPT-5, o-series)
   * reject every value but their own default, so "don't send it" has to be
   * expressible.
   */
  temperature?: number | null | undefined
  /** output token ceiling; overrides the per-request default when set */
  maxTokens?: number | undefined
  /** sent as `reasoning_effort` to OpenAI-compatible endpoints; omitted when unset */
  reasoningEffort?: ReasoningEffort | undefined
}

export interface AiProviderMeta {
  id: AiProviderId
  label: string
  models: string[]
  defaultModel: string
  keyPlaceholder: string
  needsBaseUrl?: boolean
}

/**
 * One saved OpenAI-compatible endpoint. Users keep several — a local Ollama, a
 * work gateway, a reasoning model — and switch between them from the AI
 * sidebar without opening settings.
 */
export interface AiCustomProfile extends AiProviderConfig {
  id: string
  /** user-facing name; blank falls back to the model id */
  label: string
}

export interface AiSettings {
  provider: AiProviderId
  providers: Record<AiProviderId, AiProviderConfig>
  /**
   * The saved custom endpoints, in display order. `providers.custom` stays the
   * one that requests actually use, so every request builder is unchanged; the
   * active profile is copied into it on read.
   */
  customProfiles?: AiCustomProfile[]
  /** id of the live profile; ignored while provider is not 'custom' */
  activeProfileId?: string
  /** Tavily key for web/image search and page extraction; empty falls back to the other backends */
  tavilyApiKey?: string
  /**
   * Explicit outbound proxy (`http://`, `https://`, `socks5://`, `socks4://`).
   * Empty keeps the previous behaviour: proxy env vars, then the system proxy.
   */
  proxyUrl?: string
}

/** pre-provider settings shape (single OpenAI-compatible endpoint); migrated into "custom" */
export interface LegacyAiSettings {
  baseUrl?: string
  apiKey?: string
  model?: string
}

/**
 * The AI backend as the settings dialog presents it: either the Genspark
 * account or one custom OpenAI-compatible endpoint. A flat projection of
 * AiSettings, so the UI never has to know about the provider matrix.
 */
export interface AiModelSettings {
  mode: 'genspark' | 'custom'
  /** OpenAI-compatible endpoint, e.g. https://api.deepseek.com/v1 */
  baseUrl: string
  model: string
  /** empty is allowed: local servers (Ollama, LM Studio, vLLM) accept anonymous requests */
  apiKey: string
  /** null = leave it to the model (the only setting reasoning models accept) */
  temperature: number | null
  /** null = use the app's per-request default */
  maxTokens: number | null
  /** null = don't send `reasoning_effort` */
  reasoningEffort: ReasoningEffort | null
  /** Tavily key; empty disables the Tavily backend rather than erroring */
  tavilyApiKey: string
  /** explicit proxy URL; empty means env vars / system proxy as before */
  proxyUrl: string
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
  settings: AiSettings
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
