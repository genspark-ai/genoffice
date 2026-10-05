import { ANTHROPIC_BASE_URL } from './protocols/anthropic'
import { GEMINI_BASE_URL } from './protocols/gemini'
import { AI_PROVIDERS, DEEPSEEK_V41_FLASH, GENSPARK_LLM_BASE_URLS } from './providers'
import type { AiProviderConfig, AiProviderId, AiProviderMeta } from './types'

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
 * Model families that fix sampling and reject a temperature field, on any
 * route — vendor API, the Genspark proxy, OpenRouter's vendor-prefixed ids,
 * or a mirror behind a custom base URL. Kimi K3 answers "only 1 is allowed";
 * OpenAI's GPT-5 and GPT-6 reasoning families reject any temperature other
 * than the default outright, and the o-series reasoning models (o1/o3/o4) likewise
 * only accept the default. Google's Gemini 3 docs strongly recommend keeping
 * the default temperature of 1.0 for the whole Gemini 3 family, since lower
 * values may cause looping or degraded reasoning, so our hard-coded 0.3
 * must not be sent there either.
 */
export function modelHasFixedSampling(model: string): boolean {
  return /(^|\/)(kimi-k3([^\w]|$)|gpt-[5-9]([^\w]|$)|gemini-3([^\w]|$)|o1(-mini|-preview)?([^\w]|$)|o3(-mini)?([^\w]|$)|o4-mini([^\w]|$))/i.test(
    model,
  )
}

/**
 * Model ids that reject image input even under a vision-capable provider.
 * DeepSeek V4 Pro and V4 Flash are text-only; V4.1 Flash and the -vision*
 * branches take images, so they fall through and receive screenshots.
 *
 * Ant's Ling line and Meituan's LongCat are the same shape: a text catalog
 * with one multimodal member. `Ling-3.0-flash-VL` and `LongCat-2.5-Preview`
 * (image understanding, per the 2026-09-25 LongCat change log) take images;
 * every other id on those two providers is text-only. A text-only id added to
 * either list has to be added here too — the provider flag alone would hand it
 * screenshots, which is what this function exists to prevent.
 */
export function modelLacksVision(model: string): boolean {
  // MiniMax-M2.7 remains text-only when MiniMax-M3 enables provider vision.
  if (/(^|\/)minimax-m2\.7($|-)/i.test(model)) return true
  return (
    /(^|\/)deep-?seek-v4-(?:pro(?:$|-)|flash(?!-vision))/i.test(model) ||
    /(^|\/)(?:ling-(?:3\.0-flash(?!-vl)|3\.0-tiny|2\.6-1t|2\.6-flash)|ring-2\.6-1t|longcat-2\.0(?:$|-))/i.test(
      model,
    )
  )
}

/**
 * Interleaved-thinking families whose vendors want the reasoning echoed back
 * on assistant messages: MiniMax documents that stripping it degrades
 * multi-turn tool use, and DeepSeek V4 rejects tool turns without it. Gated
 * per model because other vendors may reject the unknown field. Hunyuan joins
 * them because hy4-preview ships deep thinking on by default, so its first
 * turn already carries `reasoning_content`.
 */
export function modelEchoesReasoning(model: string): boolean {
  return /(^|\/)(minimax-m|deep-?seek-(v4|flash)|hy-?[34]([^\w]|$))/i.test(model)
}

/**
 * The direct API 400s on the versioned pool spelling we list (verified
 * 2026-09-21: GET /v1/models serves only `deepseek-flash` and `deepseek-v4-pro`).
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

/**
 * Stored provider settings are user data: a custom base URL must be a
 * bounded http(s) URL. Anything else (file:/javascript: schemes, megabyte
 * strings) would misroute gateway traffic or overflow request builders.
 *
 * A query string is kept — Azure-style bases pin `?api-version=…` and
 * gateways pin a version there — while the fragment is dropped (it is never
 * sent to the server, and leaving it on would truncate every composed
 * endpoint path). Embedded credentials are refused outright: they would end
 * up in request logs and error messages, and the api key field is the
 * supported place for them.
 */
export function normalizeBaseUrl(raw: string | undefined, fallback: string): string {
  const candidate = (raw ?? fallback).trim()
  if (candidate === '' || candidate.length > 2048) {
    throw new Error('Base URL must be a non-empty http(s) URL under 2048 characters')
  }
  let parsed: URL
  try {
    parsed = new URL(candidate)
  } catch {
    throw new Error('Base URL must be a valid http(s) URL')
  }
  if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') {
    throw new Error('Base URL must use http or https')
  }
  if (parsed.username || parsed.password) {
    throw new Error('Base URL must not embed credentials; put the key in the API key field')
  }
  parsed.hash = ''
  return parsed.toString()
}

/** Strip trailing slashes and a trailing /v1 from the path (before any query) */
function stripTrailingV1(base: string): string {
  const q = base.indexOf('?')
  const path = (q === -1 ? base : base.slice(0, q)).replace(/\/+$/, '').replace(/\/v1$/, '')
  return q === -1 ? path : `${path}${base.slice(q)}`
}

/** Append a path segment before any query string so `?api-version=…` stays last */
function appendPath(base: string, path: string): string {
  const q = base.indexOf('?')
  return q === -1 ? `${base}${path}` : `${base.slice(0, q)}${path}${base.slice(q)}`
}

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
      // No thinking override: both V4 models think by default and the agent
      // transcript round-trips the reasoning (deepseek sits on the
      // modelEchoesReasoning list). The tool-turn 400 that once forced
      // non-thinking no longer reproduces — verified against the live API
      // 2026-09-30: flash and v4-pro accept thinking+tools with and without
      // the reasoning_content echo.
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
    // the mainland TokenHub host; the international one differs only by the
    // `intl` label (tokenhub-intl.tencentcloudmaas.com), reachable by storing
    // a base URL on this provider
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://tokenhub.tencentmaas.com/v1'),
  },
  ling: {
    meta: metaOf('ling'),
    // Ling-3.0-flash-VL reads images, so the provider is vision-capable;
    // modelLacksVision() keeps the five text-only ids off screenshots
    capabilities: { auth: 'api-key', vision: true },
    // the base_url every official example uses (quickstart + OpenAI-compatible
    // reference, read 2026-10-01); /v1/models on it answers 401
    // sdk_token_not_found, so it is the live first-party host
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://api.ant-ling.com/v1'),
  },
  spark: {
    meta: metaOf('spark'),
    capabilities: { auth: 'api-key', vision: false },
    // the MaaS base from section 1.1 of the product guide (read 2026-10-01):
    // chat is POST https://maas-api.cn-huabei-1.xf-yun.com/v2/chat/completions,
    // which is this base plus the path endpointUrl() appends, so the two
    // compose back to the documented URL. The same host also serves
    // /v1/responses and /anthropic/v1/messages; we speak chat-completions
    resolveEndpoint: fixedEndpoint(
      'openai-compatible',
      'https://maas-api.cn-huabei-1.xf-yun.com/v2',
    ),
  },
  longcat: {
    meta: metaOf('longcat'),
    // 2.5-Preview reads images (2026-09-25 change log); modelLacksVision()
    // holds 2.0 back, which predates image understanding
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
      return {
        protocol: 'openai-compatible',
        baseUrl: normalizeBaseUrl(config.baseUrl, ''),
        ...(modelHasFixedSampling(config.model) ? { omitTemperature: true } : {}),
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
