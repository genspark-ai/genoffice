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

/** a stored baseUrl overrides the default endpoint (regional mirrors, e.g. api.moonshot.cn vs .ai) */
function fixedEndpoint(protocol: AiProtocol, baseUrl: string, omitTemperature?: boolean) {
  return (config: AiProviderConfig): ResolvedEndpoint => {
    const omit = omitTemperature || modelHasFixedSampling(config.model)
    return {
      protocol,
      baseUrl: config.baseUrl || baseUrl,
      ...(omit ? { omitTemperature: true } : {}),
    }
  }
}

export const AI_PROVIDER_ADAPTERS: Record<AiProviderId, ProviderAdapter> = {
  genspark: {
    meta: metaOf('genspark'),
    capabilities: { auth: 'gsk-login', vision: true },
    // The proxy exposes three protocol-specific endpoints; route by model id prefix: claude uses
    // the Anthropic protocol (preserves image input fidelity), gemini uses Gemini, rest OpenAI-compatible
    resolveEndpoint(config) {
      if (config.model.startsWith('claude')) {
        return { protocol: 'anthropic', baseUrl: GENSPARK_LLM_BASE_URLS.anthropic }
      }
      if (config.model.startsWith('gemini')) {
        return { protocol: 'gemini', baseUrl: GENSPARK_LLM_BASE_URLS.gemini }
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
    capabilities: { auth: 'api-key', vision: false },
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://api.deepseek.com/v1'),
  },
  openai: {
    meta: metaOf('openai'),
    capabilities: { auth: 'api-key', vision: true },
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://api.openai.com/v1'),
  },
  kimi: {
    meta: metaOf('kimi'),
    capabilities: { auth: 'api-key', vision: true },
    resolveEndpoint: fixedEndpoint('openai-compatible', 'https://api.moonshot.ai/v1', true),
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
