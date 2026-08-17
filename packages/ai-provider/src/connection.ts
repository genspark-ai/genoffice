import { aiFetch } from './fetch'
import { GENSPARK_LLM_BASE_URLS, OLLAMA_DEFAULT_BASE_URL } from './providers'
import type { AiProviderConfig, AiProviderId } from './types'

/**
 * Machine-readable outcome of a provider connection test. The renderer maps
 * these to localized strings; raw network errors / stack traces never leave
 * this module.
 */
export type AiConnectionStatus =
  | 'connected'
  | 'not-running'
  | 'refused'
  | 'invalid'
  | 'auth'
  | 'timeout'
  | 'unknown'

export interface AiConnectionTestResult {
  ok: boolean
  status: AiConnectionStatus
}

export interface AiConnectionTestInput {
  provider: AiProviderId
  baseUrl?: string
  apiKey?: string
  model?: string
}

const TEST_TIMEOUT_MS = 5_000

function classifyFetchError(err: unknown, provider: AiProviderId): AiConnectionStatus {
  if (err instanceof Error && err.name === 'TimeoutError') return 'timeout'
  const text = err instanceof Error ? `${err.message} ${String((err as { cause?: unknown }).cause ?? '')}` : String(err)
  const low = text.toLowerCase()
  const network =
    low.includes('econnrefused') ||
    low.includes('failed to fetch') ||
    low.includes('fetch failed') ||
    low.includes('econnreset') ||
    low.includes('network') ||
    low.includes('getaddrinfo')
  if (network) return provider === 'ollama' ? 'not-running' : 'refused'
  return 'refused'
}

function connectionStatusForResponse(status: number): AiConnectionStatus {
  if (status >= 200 && status < 300) return 'connected'
  if (status === 401 || status === 403) return 'auth'
  return 'invalid'
}

/**
 * Lightweight reachability check for the selected provider, using the least
 * invasive request that still proves auth + endpoint validity:
 * - Ollama: GET /api/tags (the native tags endpoint, no auth)
 * - OpenAI-compatible (openai / deepseek / custom / genspark): GET /models
 * - Gemini: GET /models (x-goog-api-key)
 * - Anthropic: GET /v1/models (x-api-key)
 * Never returns stack traces or raw bodies — only a friendly status code.
 */
export async function testProviderConnection(
  provider: AiProviderId,
  config: AiProviderConfig,
): Promise<AiConnectionTestResult> {
  const headers: Record<string, string> = { 'Content-Type': 'application/json' }
  let url: string | null = null

  switch (provider) {
    case 'ollama': {
      const root = (config.baseUrl || OLLAMA_DEFAULT_BASE_URL)
        .replace(/\/v1\/?$/, '')
        .replace(/\/+$/, '')
      url = `${root}/api/tags`
      break
    }
    case 'openai':
    case 'deepseek':
    case 'custom': {
      const fallback =
        provider === 'openai' ? 'https://api.openai.com/v1' : provider === 'deepseek' ? 'https://api.deepseek.com/v1' : ''
      const base = config.baseUrl || fallback
      if (!base) return { ok: false, status: 'invalid' }
      url = `${base.replace(/\/+$/, '')}/models`
      if (config.apiKey) headers.Authorization = `Bearer ${config.apiKey}`
      break
    }
    case 'gemini': {
      const base = config.baseUrl || 'https://generativelanguage.googleapis.com/v1beta'
      url = `${base.replace(/\/+$/, '')}/models`
      if (config.apiKey) headers['x-goog-api-key'] = config.apiKey
      break
    }
    case 'anthropic': {
      const base = config.baseUrl || 'https://api.anthropic.com'
      url = `${base.replace(/\/+$/, '')}/v1/models`
      if (config.apiKey) headers['x-api-key'] = config.apiKey
      break
    }
    case 'genspark': {
      // the proxy's OpenAI-compatible route; the caller injects the gsk key
      url = `${GENSPARK_LLM_BASE_URLS.openai.replace(/\/+$/, '')}/models`
      if (config.apiKey) headers.Authorization = `Bearer ${config.apiKey}`
      break
    }
    default:
      return { ok: false, status: 'unknown' }
  }

  try {
    const resp = await aiFetch(url, {
      method: 'GET',
      signal: AbortSignal.timeout(TEST_TIMEOUT_MS),
      headers,
    })
    return { ok: resp.ok, status: connectionStatusForResponse(resp.status) }
  } catch (err) {
    return { ok: false, status: classifyFetchError(err, provider) }
  }
}
