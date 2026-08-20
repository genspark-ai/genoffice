import type { AgentMessage, AgentToolCall, AgentToolDef } from '@genoffice/agent-core'

export type AiProviderId =
  | 'genspark'
  | 'anthropic'
  | 'gemini'
  | 'deepseek'
  | 'openai'
  | 'kimi'
  | 'glm'
  | 'qwen'
  | 'doubao'
  | 'minimax'
  | 'xai'
  | 'mistral'
  | 'openrouter'
  | 'custom'

/** Genspark account status (gsk login state; the sole auth source for AI features) */
export interface GenSparkAccountStatus {
  loggedIn: boolean
  email?: string
}

export interface AiProviderConfig {
  apiKey: string
  model: string
  /** required for custom; for other direct providers it overrides the default endpoint (regional mirrors) */
  baseUrl?: string | undefined
}

export interface AiProviderMeta {
  id: AiProviderId
  label: string
  models: string[]
  defaultModel: string
  keyPlaceholder: string
  needsBaseUrl?: boolean
}

export interface AiSettings {
  provider: AiProviderId
  providers: Record<AiProviderId, AiProviderConfig>
  /**
   * Genspark cloud tools (web/image search via gsk, image generation, media
   * analysis). Default true; false makes tools skip the gsk backend entirely
   * (search falls back to free sources, gsk-only tools are unavailable).
   * Only meaningful while signed in — signed out, the gsk backend is
   * unavailable regardless.
   */
  gskToolsEnabled?: boolean
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
  /** machine-readable error cause ('timeout', exhausted 'credits', 'network' connectivity failure); lets the renderer localize the message */
  errorCode?: 'timeout' | 'credits' | 'network'
  /** normalized stop reason carried on 'done' ('max_tokens' = output cut off by the token limit) */
  stopReason?: string
}
