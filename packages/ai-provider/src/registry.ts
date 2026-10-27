import { ANTHROPIC_BASE_URL } from './protocols/anthropic'
import { GEMINI_BASE_URL } from './protocols/gemini'
import { AI_PROVIDERS, DEEPSEEK_V41_FLASH, GENSPARK_LLM_BASE_URLS } from './providers'
import {
  customProtocolBaseUrl,
  normalizeCustomProfile,
  resolveCustomProfile,
} from './custom-profile'
import { appendPath, normalizeBaseUrl, stripTrailingV1 } from './url-compose'
import { modelHasFixedSampling, modelLacksVision } from './model-heuristics'
import type { AiProviderConfig, AiProviderId, AiProviderMeta } from './types'

/** The shipped model-family guesses live in ./model-heuristics; re-exported for callers. */
export { normalizeBaseUrl } from './url-compose'
export { modelEchoesReasoning, modelHasFixedSampling, modelLacksVision } from './model-heuristics'

/** Wire protocols every provider maps onto, including the official Codex app-server bridge. */
export type AiProtocol = 'anthropic' | 'gemini' | 'openai-compatible' | 'codex-app-server'

export interface ProviderCapabilities {
  /** How the provider authenticates: app login, user key, or the Codex CLI's existing login. */
  auth: 'gsk-login' | 'api-key' | 'codex-chatgpt'
  /** chat models accept image input (declarative; for custom endpoints it is assumed, not known) */
  vision: boolean
}

export interface ResolvedEndpoint {
  protocol: AiProtocol
  baseUrl: string
  /** the endpoint fixes its sampling and rejects a temperature field (Kimi K3: "only 1 is allowed") */
  omitTemperature?: boolean
  /** the endpoint wants the output cap as OpenAI's renamed `max_completion_tokens` (GPT-5.x/o-series 400 on `max_tokens`) */
  useMaxCompletionTokens?: boolean
  /** vendor-specific request fields merged into the chat-completions body */
  bodyExtras?: Record<string, unknown>
  /** id to put on the wire when the vendor spells the configured model differently */
  model?: string
  /** the endpoint wants `reasoning_content` echoed back on assistant tool turns */
  echoReasoning?: boolean
  /** static headers the user's endpoint needs (tenant id, gateway routing key) */
  headers?: Record<string, string>
  /** stream idle budget override for this endpoint (slow local servers need minutes) */
  idleTimeoutMs?: number
  /** the configured model's declared output ceiling, honored over the global cap */
  maxTokens?: number
  /** images may be sent to this endpoint (custom: from the profile, else the heuristic) */
  vision?: boolean
}

export interface ProviderAdapter {
  meta: AiProviderMeta
  capabilities: ProviderCapabilities
  /** pick the wire protocol and base URL for one request (may depend on the configured model) */
  resolveEndpoint(config: AiProviderConfig): ResolvedEndpoint
}

function metaOf(id: AiProviderId): AiProviderMeta {
  return AI_PROVIDERS.find((m) => m.id === id)!
}

/**
 * The pool spelling is not a vendor id: the direct API 400s on it (verified
 * 2026-09-21 — GET /v1/models serves only `deepseek-flash` and
 * `deepseek-v4-pro`). Settings read from disk are rewritten by RETIRED_MODELS,
 * but in-memory state bypasses that, so the adapter still maps the alias here.
 */
const DEEPSEEK_WIRE_IDS: Record<string, string> = { [DEEPSEEK_V41_FLASH]: 'deepseek-flash' }

/**
 * OpenCode Zen / Go (opencode.ai) are protocol passthrough gateways: each
 * model is served on exactly one vendor protocol and the other paths answer
 * 500 (verified 2026-09-03 against the public free tier), so the route is
 * picked per model id the way OpenCode's own client does (models.dev
 * `provider.npm`). The two tiers route the same vendor differently — MiniMax
 * is chat-completions on Zen but Anthropic Messages on Go — hence one table
 * each. Every Kimi id omits temperature, mirroring the direct Kimi adapter.
 */
const OPENCODE_GATEWAY_ROOTS = {
  zen: 'https://opencode.ai/zen',
  go: 'https://opencode.ai/zen/go',
} as const

function opencodeEndpoint(
  root: string,
  routes: { anthropic: RegExp; gemini?: RegExp },
): (config: AiProviderConfig) => ResolvedEndpoint {
  return (config) => {
    // a stored base URL replaces the gateway root; the documented `/v1` API base is tolerated
    const base = stripTrailingV1(normalizeBaseUrl(config.baseUrl, root))
    const model = config.model ?? ''
    const omit =
      model !== '' && (modelHasFixedSampling(model) || model.toLowerCase().startsWith('kimi-'))
    const sampling = omit ? { omitTemperature: true as const } : {}
    if (routes.anthropic.test(model)) return { protocol: 'anthropic', baseUrl: base, ...sampling }
    if (routes.gemini?.test(model)) {
      return { protocol: 'gemini', baseUrl: appendPath(base, '/v1'), ...sampling }
    }
    return { protocol: 'openai-compatible', baseUrl: appendPath(base, '/v1'), ...sampling }
  }
}

/** a stored baseUrl overrides the default endpoint (regional mirrors, e.g. api.moonshot.cn vs .ai) */
function fixedEndpoint(
  protocol: AiProtocol,
  baseUrl: string,
  extras?: {
    omitTemperature?: boolean
    useMaxCompletionTokens?: boolean
    bodyExtras?: Record<string, unknown>
  },
) {
  return (config: AiProviderConfig): ResolvedEndpoint => {
    const omit = extras?.omitTemperature || modelHasFixedSampling(config.model)
    return {
      protocol,
      baseUrl: config.baseUrl ? normalizeBaseUrl(config.baseUrl, baseUrl) : baseUrl,
      ...(omit ? { omitTemperature: true } : {}),
      ...(extras?.useMaxCompletionTokens ? { useMaxCompletionTokens: true } : {}),
      ...(extras?.bodyExtras ? { bodyExtras: extras.bodyExtras } : {}),
    }
  }
}

export const AI_PROVIDER_ADAPTERS: Record<AiProviderId, ProviderAdapter> = {
  genspark: {
    meta: metaOf('genspark'),
    capabilities: { auth: 'gsk-login', vision: true },
    // Route by model id prefix: claude uses the Anthropic protocol (preserves image
    // input fidelity), the rest OpenAI-compatible. The proxy's gemini endpoint was
    // removed server-side (405 as of 2026-08-31) along with its gemini models.
    resolveEndpoint(config) {
      if (config.model.startsWith('claude')) {
        return { protocol: 'anthropic', baseUrl: GENSPARK_LLM_BASE_URLS.anthropic }
      }
      return {
        protocol: 'openai-compatible',
        baseUrl: GENSPARK_LLM_BASE_URLS.openai,
        ...(modelHasFixedSampling(config.model) ? { omitTemperature: true } : {}),
      }
    },
  },
  codex: {
    meta: metaOf('codex'),
    capabilities: { auth: 'codex-chatgpt', vision: true },
    resolveEndpoint() {
      return { protocol: 'codex-app-server', baseUrl: '' }
    },
  },
  anthropic: {
    meta: metaOf('anthropic'),
    capabilities: { auth: 'api-key', vision: true },
    resolveEndpoint: fixedEndpoint('anthropic', ANTHROPIC_BASE_URL),
  },
  gemini: {
    meta: metaOf('gemini'),
    capabilities: { auth: 'api-key', vision: true },
    resolveEndpoint: fixedEndpoint('gemini', GEMINI_BASE_URL),
  },
  deepseek: {
    meta: metaOf('deepseek'),
    capabilities: { auth: 'api-key', vision: true },
    resolveEndpoint(config) {
      const wire = DEEPSEEK_WIRE_IDS[config.model]
      // thinking stays on (vendor default); the transcript echoes reasoning_content for tool turns
      return {
        ...fixedEndpoint('openai-compatible', 'https://api.deepseek.com/v1')(config),
        ...(wire ? { model: wire } : {}),
      }
    },
  },
  openai: {
    meta: metaOf('openai'),
    capabilities: { auth: 'api-key', vision: true },
    // every current OpenAI model accepts the renamed field, so it is safe endpoint-wide;
    // other openai-compatible vendors (and the LiteLLM-backed Genspark proxy) still expect `max_tokens`
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://api.openai.com/v1', {
      useMaxCompletionTokens: true,
    }),
  },
  kimi: {
    meta: metaOf('kimi'),
    capabilities: { auth: 'api-key', vision: true },
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://api.moonshot.ai/v1', {
      omitTemperature: true,
    }),
  },
  glm: {
    meta: metaOf('glm'),
    capabilities: { auth: 'api-key', vision: false },
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://open.bigmodel.cn/api/paas/v4'),
  },
  qwen: {
    meta: metaOf('qwen'),
    capabilities: { auth: 'api-key', vision: false },
    resolveEndpoint: fixedEndpoint(
      'openai-compatible',
      'https://dashscope.aliyuncs.com/compatible-mode/v1',
    ),
  },
  doubao: {
    meta: metaOf('doubao'),
    capabilities: { auth: 'api-key', vision: true },
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://ark.cn-beijing.volces.com/api/v3'),
  },
  mimo: {
    meta: metaOf('mimo'),
    // the V2.6 series is omni-modal: text, image, video and audio in, text out
    capabilities: { auth: 'api-key', vision: true },
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://api.xiaomimimo.com/v1'),
  },
  hunyuan: {
    meta: metaOf('hunyuan'),
    // conservative: the chat models are documented for text first, so we do not
    // hand them screenshots until a model card says otherwise
    capabilities: { auth: 'api-key', vision: false },
    // mainland TokenHub host; the intl one (tokenhub-intl.tencentcloudmaas.com)
    // is reachable by storing a base URL on this provider
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://tokenhub.tencentcloudmaas.com/v1'),
  },
  ling: {
    meta: metaOf('ling'),
    capabilities: { auth: 'api-key', vision: true },
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://api.ant-ling.com/v1'),
  },
  spark: {
    meta: metaOf('spark'),
    capabilities: { auth: 'api-key', vision: false },
    // endpointUrl() appends chat/completions to this documented /v2 base
    resolveEndpoint: fixedEndpoint(
      'openai-compatible',
      'https://maas-api.cn-huabei-1.xf-yun.com/v2',
    ),
  },
  longcat: {
    meta: metaOf('longcat'),
    capabilities: { auth: 'api-key', vision: true },
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://api.longcat.chat/openai/v1'),
  },
  minimax: {
    meta: metaOf('minimax'),
    capabilities: { auth: 'api-key', vision: true },
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://api.minimax.io/v1'),
  },
  xai: {
    meta: metaOf('xai'),
    capabilities: { auth: 'api-key', vision: true },
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://api.x.ai/v1'),
  },
  mistral: {
    meta: metaOf('mistral'),
    capabilities: { auth: 'api-key', vision: false },
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://api.mistral.ai/v1'),
  },
  openrouter: {
    meta: metaOf('openrouter'),
    capabilities: { auth: 'api-key', vision: true },
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://openrouter.ai/api/v1'),
  },
  requesty: {
    meta: metaOf('requesty'),
    capabilities: { auth: 'api-key', vision: true },
    // a stored base URL selects a regional router (https://router.eu.requesty.ai/v1 for the EU)
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://router.requesty.ai/v1'),
  },
  opper: {
    meta: metaOf('opper'),
    capabilities: { auth: 'api-key', vision: true },
    // one chat-completions endpoint for every pool and vendor route; the model id picks it
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://api.opper.ai/v3/compat'),
  },
  cheaperinference: {
    meta: metaOf('cheaperinference'),
    capabilities: { auth: 'api-key', vision: true },
    // one chat-completions endpoint for every model; the model id picks the lab
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://api.cheaperinference.com/v1'),
  },
  atlascloud: {
    meta: metaOf('atlascloud'),
    capabilities: { auth: 'api-key', vision: true },
    // one chat-completions endpoint for every lab; the namespaced model id picks it
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://api.atlascloud.ai/v1'),
  },
  runcrate: {
    meta: metaOf('runcrate'),
    capabilities: { auth: 'api-key', vision: true },
    // one chat-completions endpoint for every lab; the namespaced model id picks it
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://api.runcrate.ai/v1'),
  },
  'opencode-zen': {
    meta: metaOf('opencode-zen'),
    capabilities: { auth: 'api-key', vision: true },
    // Claude and Qwen ride /v1/messages, Gemini its native generateContent path
    resolveEndpoint: opencodeEndpoint(OPENCODE_GATEWAY_ROOTS.zen, {
      anthropic: /^(claude-|qwen)/,
      gemini: /^gemini-/,
    }),
  },
  'opencode-go': {
    meta: metaOf('opencode-go'),
    capabilities: { auth: 'api-key', vision: true },
    // MiniMax and Qwen 3.8 Flash ride /v1/messages; the rest is chat-completions
    // (the Go docs table lists every Qwen on Messages, but the client config
    // OpenCode ships routes only 3.8 Flash there — follow the running client)
    resolveEndpoint: opencodeEndpoint(OPENCODE_GATEWAY_ROOTS.go, {
      anthropic: /^(minimax-|qwen3\.8-flash$)/,
    }),
  },
  custom: {
    meta: metaOf('custom'),
    capabilities: { auth: 'api-key', vision: true },
    resolveEndpoint(config) {
      if (!config.baseUrl) throw new Error('A custom provider requires a Base URL')
      const profile = normalizeCustomProfile(config.profile)
      // Resolve against the id the user configured: an alias rewrites the wire
      // value but must not hide the declaration that produced it.
      const resolved = resolveCustomProfile(profile, config.model ?? '')
      return {
        protocol: resolved.protocol,
        baseUrl: customProtocolBaseUrl(resolved.protocol, config.baseUrl),
        ...(resolved.omitTemperature ? { omitTemperature: true } : {}),
        ...(resolved.useMaxCompletionTokens ? { useMaxCompletionTokens: true } : {}),
        ...(resolved.echoReasoning !== undefined ? { echoReasoning: resolved.echoReasoning } : {}),
        ...(resolved.bodyExtras ? { bodyExtras: resolved.bodyExtras } : {}),
        ...(resolved.wireModel ? { model: resolved.wireModel } : {}),
        ...(resolved.headers ? { headers: resolved.headers } : {}),
        ...(resolved.idleTimeoutMs !== undefined ? { idleTimeoutMs: resolved.idleTimeoutMs } : {}),
        ...(resolved.maxTokens !== undefined ? { maxTokens: resolved.maxTokens } : {}),
        ...(resolved.vision !== undefined ? { vision: resolved.vision } : {}),
      }
    },
  },
}

/** Throws on ids not in the registry — settings files are user data and can carry anything. */
export function getProviderAdapter(provider: AiProviderId): ProviderAdapter {
  const adapter = AI_PROVIDER_ADAPTERS[provider]
  if (!adapter) throw new Error(`Unknown provider: ${provider}`)
  return adapter
}

/**
 * Whether images may be sent to the active provider's configured model.
 *
 * A custom endpoint answers from its declared profile when it carries one
 * (the whole point: the user knows what their server takes, the shipped
 * heuristics do not); otherwise the provider capability plus the text-only
 * model list decide, exactly as before. Anything the registry rejects — an
 * unknown provider id in a hand-edited file — is treated as text-only, the
 * same conservative answer callers made before profiles existed.
 */
export function endpointSupportsVision(settings: {
  provider: AiProviderId
  providers?: Partial<Record<AiProviderId, AiProviderConfig>> | undefined
}): boolean {
  try {
    const adapter = getProviderAdapter(settings.provider)
    if (!adapter.capabilities.vision) return false
    const config = settings.providers?.[settings.provider]
    if (!config) return false
    return adapter.resolveEndpoint(config).vision ?? !modelLacksVision(config.model)
  } catch {
    return false
  }
}
