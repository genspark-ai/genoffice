import { ANTHROPIC_BASE_URL } from './protocols/anthropic'
import { GEMINI_BASE_URL } from './protocols/gemini'
import { AI_PROVIDERS, GENSPARK_LLM_BASE_URLS } from './providers'
import type { AiProviderConfig, AiProviderId, AiProviderMeta } from './types'

/** The three wire protocols every provider maps onto. */
export type AiProtocol = 'anthropic' | 'gemini' | 'openai-compatible'

export interface ProviderCapabilities {
  /** 'gsk-login': the Genspark session key is injected by the main process; 'api-key': the user supplies their own key */
  auth: 'gsk-login' | 'api-key'
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
 * OpenAI's GPT-5 reasoning family rejects any temperature other than the
 * default outright.
 */
export function modelHasFixedSampling(model: string): boolean {
  return /(^|\/)(kimi-k3|gpt-5)/.test(model)
}

/**
 * Model ids that reject image input even under a vision-capable provider.
 * DeepSeek V4 Pro and Flash are text-only; their -vision* branches are
 * excluded so the direct Vision Exp model can receive screenshots.
 */
export function modelLacksVision(model: string): boolean {
  return /(^|\/)deep-?seek-v4-(?:pro(?:$|-)|flash(?!-vision))/.test(model)
}

/**
 * Interleaved-thinking families whose vendors want the reasoning echoed back
 * on assistant messages: MiniMax documents that stripping it degrades
 * multi-turn tool use, and DeepSeek V4 rejects tool turns without it. Gated
 * per model because other vendors may reject the unknown field.
 */
export function modelEchoesReasoning(model: string): boolean {
  return /(^|\/)(minimax-m|deep-?seek-v4)/i.test(model)
}

/**
 * DeepSeek V4 thinks by default, and once a request carries `tools` the API
 * rejects (400) every later turn whose assistant messages don't echo back the
 * `reasoning_content` it produced. Our OpenAI-compatible transcript has no
 * field to carry that, so the agent loop would die right after its first tool
 * call. Pin the models to non-thinking mode — what the retired deepseek-chat
 * alias did — until the transcript can round-trip reasoning.
 */
const DEEPSEEK_NON_THINKING = { thinking: { type: 'disabled' } }

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
      baseUrl: config.baseUrl || baseUrl,
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
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://api.deepseek.com/v1', {
      bodyExtras: DEEPSEEK_NON_THINKING,
    }),
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
  minimax: {
    meta: metaOf('minimax'),
    capabilities: { auth: 'api-key', vision: false },
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
  custom: {
    meta: metaOf('custom'),
    capabilities: { auth: 'api-key', vision: true },
    resolveEndpoint(config) {
      if (!config.baseUrl) throw new Error('A custom provider requires a Base URL')
      return {
        protocol: 'openai-compatible',
        baseUrl: config.baseUrl,
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
