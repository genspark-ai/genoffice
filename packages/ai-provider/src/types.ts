import type { AgentMessage, AgentToolCall, AgentToolDef } from '@genoffice/agent-core'

export type AiProviderId =
  'genspark' | 'anthropic' | 'gemini' | 'deepseek' | 'openai' | 'openai-codex' | 'custom'

/** Genspark account status (gsk login state; the sole auth source for AI features) */
export interface GenSparkAccountStatus {
  loggedIn: boolean
  email?: string
}

export interface AiProviderConfig {
  apiKey: string
  model: string
  /** Codex-only Responses reasoning setting; `none` omits the request field. */
  reasoningEffort?: CodexReasoningEffort
  /** only used by the custom (OpenAI-compatible) provider */
  baseUrl?: string | undefined
}

export type CodexReasoningEffort =
  'none' | 'minimal' | 'low' | 'medium' | 'high' | 'xhigh' | 'max' | 'ultra'

/** Renderer-safe, account-specific Codex model capabilities. */
export interface CodexModelCapability {
  id: string
  reasoningEfforts: CodexReasoningEffort[]
}

export interface CodexCapabilities {
  models: CodexModelCapability[]
}

export interface AiProviderMeta {
  id: AiProviderId
  label: string
  models: string[]
  defaultModel: string
  keyPlaceholder: string
  /** false when this provider receives main-process account credentials instead of a settings API key */
  requiresApiKey?: boolean
  needsBaseUrl?: boolean
}

/**
 * Account credentials injected by the main process for one Codex request.
 * Never persist this in AiSettings or forward it over renderer IPC.
 */
export interface CodexAuthContext {
  accessToken: string
  accountId: string
  expiresAt: number
}

/** Narrow provider boundary; app code owns prompts, history, and tool execution. */
export interface CodexAdapterRequest {
  auth: CodexAuthContext
  instructions: string
  messages: AgentMessage[]
  tools: AgentToolDef[]
  model: string
  reasoningEffort?: CodexReasoningEffort
  signal: AbortSignal
  onDelta: (text: string) => void
  onToolCall: (call: AgentToolCall) => void
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
