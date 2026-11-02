import type { AgentMessage, AgentToolCall, AgentToolDef } from '@genoffice/agent-core'

export type AiProviderId =
  | 'genspark'
  | 'codex'
  | 'anthropic'
  | 'gemini'
  | 'deepseek'
  | 'openai'
  | 'kimi'
  | 'glm'
  | 'qwen'
  | 'doubao'
  | 'mimo'
  | 'hunyuan'
  | 'ling'
  | 'spark'
  | 'longcat'
  | 'minimax'
  | 'xai'
  | 'mistral'
  | 'openrouter'
  | 'requesty'
  | 'opper'
  | 'cheaperinference'
  | 'atlascloud'
  | 'runcrate'
  | 'cloudgpu'
  | 'opencode-zen'
  | 'opencode-go'
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
  /** optional Codex CLI override; empty means auto-detect the current authenticated install */
  cliPath?: string | undefined
  /** the declarative profile of the active user-hosted endpoint (custom only) */
  profile?: AiCustomProfile | undefined
}

/**
 * Wire protocols a user-hosted endpoint can speak. The three the package
 * already implements, so a self-hosted gateway that ships an Anthropic- or
 * Gemini-shaped route becomes reachable without any code change.
 */
export type AiCustomProtocol = 'openai-compatible' | 'anthropic' | 'gemini'

/**
 * Vendor quirks an endpoint can declare instead of leaving the built-in
 * model-name heuristics to guess. Absent means "fall through to the
 * heuristic", so an endpoint that says nothing behaves exactly as before.
 */
export interface AiCustomCompat {
  /** the endpoint 400s on a temperature field (fixed-sampling models behind a proxy) */
  omitTemperature?: boolean | undefined
  /** the endpoint wants `reasoning_content` echoed back on assistant tool turns */
  echoReasoning?: boolean | undefined
  /** false declares a text-only endpoint so screenshots are never sent to it */
  vision?: boolean | undefined
  /** send the output cap as `max_completion_tokens` (OpenAI renamed field) */
  useMaxCompletionTokens?: boolean | undefined
}

/**
 * One model declared by the user for their endpoint, so a vendor's newest
 * release — which no shipped heuristic knows yet — carries real facts instead
 * of a guess. Fields left unset fall through to the endpoint compat, then the
 * heuristics; `alias` maps the listed id to the spelling the wire expects.
 */
export interface AiCustomModelProfile {
  id: string
  /** the id to put on the wire when the endpoint spells this model differently */
  alias?: string | undefined
  /** display name for the picker (a vendor's marketing spelling of the id) */
  name?: string | undefined
  /** declared capabilities; absent = inherit the endpoint compat, then heuristics */
  vision?: boolean | undefined
  omitTemperature?: boolean | undefined
  echoReasoning?: boolean | undefined
  useMaxCompletionTokens?: boolean | undefined
  /** this model's per-turn output ceiling; a server 400s anything above it */
  maxTokens?: number | undefined
  /** the model's context window, when the endpoint publishes one */
  contextWindow?: number | undefined
}

/**
 * The declarative half of a custom endpoint. It mirrors into
 * `providers.custom.profile` so the protocol layer reads one config object,
 * and it is what makes a self-hosted or regional gateway fully describable
 * without touching code: protocol, headers, credential indirection, timeouts,
 * vendor switches, and a per-model capability list.
 */
export interface AiCustomProfile {
  /** wire protocol; absent means 'openai-compatible' (the pre-profile shape) */
  api?: AiCustomProtocol | undefined
  /**
   * Name of an environment variable the main process resolves to the key at
   * request time. Lets a settings file carry no secret, so it can be synced or
   * shared; an inline `apiKey` always wins.
   */
  apiKeyEnv?: string | undefined
  /** static request headers (tenant ids, gateway routing keys) */
  headers?: Record<string, string> | undefined
  /** stream idle budget for slow local servers (quantized models take minutes to first token) */
  idleTimeoutMs?: number | undefined
  /** endpoint-wide vendor switches, overridden per model by `modelProfiles` */
  compat?: AiCustomCompat | undefined
  /** vendor-specific fields merged into the request body (vLLM `chat_template_kwargs`, …) */
  bodyExtras?: Record<string, unknown> | undefined
  /** per-model declarations, matched to the configured id on request */
  modelProfiles?: AiCustomModelProfile[] | undefined
  /**
   * Capability facts the last `/models` probe observed, kept beside the id list
   * as a cache. They answer for a model the user has not declared — the whole
   * point, since a gateway knows things no shipped heuristic does — and they
   * are replaced wholesale by the next probe, so a model the server stops
   * serving stops being described. The user's `modelProfiles` always outrank
   * them.
   */
  probedProfiles?: AiCustomModelProfile[] | undefined
}

/** The profile fields, spread flat so a hand-edited settings file stays readable */
export interface AiCustomEndpoint extends AiCustomProfile {
  id: string
  name: string
  baseUrl: string
  apiKey: string
  model: string
  /** ids the server advertised on the last `/models` probe, so the composer chip can list them offline */
  models?: string[] | undefined
}

/** Live picker data returned by Codex app-server's model/list method. */
export interface CodexModelCatalog {
  models: string[]
  defaultModel: string
  /**
   * Capability declarations derived from a custom endpoint's own `/models`
   * metadata (vision, output ceiling, context window). Absent for Codex and
   * for endpoints whose listing carries no such fields — the settings pane
   * folds them into the endpoint's model declarations on save.
   */
  profiles?: AiCustomModelProfile[] | undefined
}

export interface AiProviderMeta {
  id: AiProviderId
  label: string
  models: string[]
  defaultModel: string
  keyPlaceholder: string
  needsBaseUrl?: boolean
  needsCliPath?: boolean
  /** the settings page refreshes `models` from the vendor's own /models route; the shipped list is the offline fallback */
  liveModels?: boolean
}

/** Image generation / media analysis backends (separate from the chat provider) */
export type AiMediaProviderId =
  | 'genspark'
  | 'openai'
  | 'gemini'
  | 'doubao'
  | 'glm'
  | 'xai'
  | 'qwen'
  | 'minimax'
  | 'deepseek'
  | 'custom'

/** wire shape of the image endpoint */
export type AiImageProtocol = 'openai-images' | 'gemini' | 'dashscope' | 'minimax'
/** wire shape of the understanding endpoint */
export type AiAnalysisProtocol = 'openai-chat' | 'gemini'

export interface AiMediaProviderConfig {
  apiKey: string
  /** required for custom; for the others it overrides the official endpoint (regional mirrors) */
  baseUrl?: string | undefined
  /** image generation model (empty = the provider default) */
  imageModel: string
  /** image/video understanding model (empty = the provider default) */
  analysisModel: string
}

export interface AiMediaProviderMeta {
  id: AiMediaProviderId
  label: string
  /** one-line English blurb shown on the provider card */
  description: string
  keyPlaceholder: string
  needsBaseUrl?: boolean
  /** '' for genspark (gsk login) and custom (user-supplied) */
  defaultBaseUrl: string
  /** absent = the provider does not generate images */
  imageProtocol?: AiImageProtocol
  imageModels: string[]
  defaultImageModel: string
  /** absent = the provider does not analyze media */
  analysisProtocol?: AiAnalysisProtocol
  analysisModels: string[]
  defaultAnalysisModel: string
  /** the analysis model accepts video input (Gemini natively; OpenAI-compatible vendors via a video_url part) */
  videoAnalysis: boolean
}

export interface AiMediaSettings {
  /** provider behind generate_image */
  imageProvider: AiMediaProviderId
  /** provider behind analyze_media for images */
  analysisProvider: AiMediaProviderId
  /** provider behind analyze_media when the input has video/audio (only video-capable vendors qualify) */
  videoAnalysisProvider: AiMediaProviderId
  providers: Record<AiMediaProviderId, AiMediaProviderConfig>
  /** pre-catalog shape (one provider for both); migrated by resolveAiMediaSettings */
  provider?: AiMediaProviderId | undefined
}

/** web/image search backends; Parallel and Keenable take a user key or run free and keyless */
export type AiSearchProviderId =
  'genspark' | 'serper' | 'serply' | 'tavily' | 'parallel' | 'exa' | 'firecrawl' | 'keenable'

export interface AiSearchProviderMeta {
  id: AiSearchProviderId
  label: string
  keyPlaceholder: string
  /** the backend also serves image search (otherwise image search falls back to free sources) */
  imageSearch: boolean
}

export interface AiSearchSettings {
  provider: AiSearchProviderId
  providers: Record<Exclude<AiSearchProviderId, 'genspark'>, { apiKey: string }>
}

export interface AiSettings {
  provider: AiProviderId
  providers: Record<AiProviderId, AiProviderConfig>
  /** every saved custom endpoint; absent on files written before the list existed */
  customEndpoints?: AiCustomEndpoint[] | undefined
  /** id of the endpoint `providers.custom` mirrors */
  customEndpoint?: string | undefined
  /**
   * Provider for generate_image / analyze_media. Absent (pre-media settings
   * files) means Genspark, i.e. the gsk login + gskToolsEnabled gate.
   */
  media?: AiMediaSettings | undefined
  /** web/image search backend; absent means Genspark (gsk when signed in, then the free chain) */
  search?: AiSearchSettings | undefined
  /**
   * Genspark cloud tools (web/image search via gsk, image generation, media
   * analysis). Default true; false makes tools skip the gsk backend entirely
   * (search falls back to free sources, gsk-only tools are unavailable).
   * Only meaningful while signed in — signed out, the gsk backend is
   * unavailable regardless.
   */
  gskToolsEnabled?: boolean
  /**
   * Output-token cap for ONE model turn of agent runs (default
   * DEFAULT_MAX_OUTPUT_TOKENS). Reasoning models bill their thinking against
   * this same budget, so a heavy edit turn can consume all of it and close with
   * finish_reason=length and no prose at all — raising it is the user's lever
   * (absent = the default, so pre-existing settings files keep working).
   */
  maxOutputTokens?: number | undefined
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
  /** Stable renderer transport id used to retain native provider sessions. */
  sessionId?: string
  settings: AiSettings
  system: string
  messages: AgentMessage[]
  tools?: AgentToolDef[]
  maxTokens?: number
}

/** Token usage a provider reported for one model turn. Both fields are best-effort:
 * some providers (and the Genspark proxy) never report usage, and consumers must
 * treat an absent figure as "unknown", never as zero. */
export interface AiTokenUsage {
  inputTokens?: number
  outputTokens?: number
}

export interface AiStreamChunk {
  requestId: string
  /** 'ping' = wire-level keepalive so the renderer can tell a live stream from a dead one;
   * 'reasoning' = model thinking delta (text carries it), stored for interleaved-thinking echo */
  type: 'delta' | 'reasoning' | 'tool-call' | 'usage' | 'done' | 'error' | 'ping'
  text?: string
  /** complete parsed tool call (emitted once its arguments finish streaming) */
  toolCall?: AgentToolCall
  error?: string
  /** machine-readable error cause ('timeout', exhausted 'credits', 'network' connectivity failure, 'overloaded' capacity/rate limit, 'region' vendor refuses the request origin); lets the renderer localize the message */
  errorCode?: 'timeout' | 'credits' | 'network' | 'overloaded' | 'region'
  /** normalized stop reason carried on 'done' ('max_tokens' = output cut off by the token limit) */
  stopReason?: string
  /** tokens the provider billed for this turn, carried on 'usage' (always sent before 'done') */
  usage?: AiTokenUsage
}
