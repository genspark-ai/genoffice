import { chatAnthropic } from './protocols/anthropic'
import { chatGemini } from './protocols/gemini'
import { chatOpenAiCompatible } from './protocols/openai-compatible'
import { chatOpenAiResponses } from './protocols/openai-responses'
import { ResponseBodyTooLargeError } from './protocols/shared'
import { chatCodexAppServer } from './codex-app-server'
import { getProviderAdapter, type ResolvedEndpoint } from './registry'
import type { AiChatResponse, AiProviderConfig, AiProviderId } from './types'
import { AI_CHAT_RESPONSE_TIMEOUT_MS, createStreamWatchdog } from './watchdog'

/** route a one-shot (non-streaming, non-tool-calling) chat call by provider id */
export async function chatForProvider(
  provider: AiProviderId,
  config: AiProviderConfig,
  system: string,
  user: string,
  signal?: AbortSignal,
): Promise<AiChatResponse> {
  // Resolve the endpoint before the watchdog: a config error (unknown
  // provider, missing base URL) reports as a failed reply, not a rejection,
  // and an endpoint's own idle budget must reach the watchdog before the
  // request starts.
  let endpoint: ResolvedEndpoint
  try {
    endpoint = getProviderAdapter(provider).resolveEndpoint(config)
  } catch (e) {
    return { ok: false as const, error: e instanceof Error ? e.message : String(e) }
  }
  // non-streaming: the server generates the full answer before the headers arrive,
  // so the connect phase gets the long budget; the body read then gets the idle budget
  const wd = createStreamWatchdog(signal, AI_CHAT_RESPONSE_TIMEOUT_MS, endpoint.idleTimeoutMs)
  const result = wd.guard(() => {
    if (endpoint.model) config = { ...config, model: endpoint.model }
    switch (endpoint.protocol) {
      case 'codex-app-server':
        return chatCodexAppServer(config, system, user, wd.signal)
      case 'anthropic':
        return chatAnthropic(wd, config, system, user, endpoint.baseUrl, {
          headers: endpoint.headers,
        })
      case 'gemini':
        return chatGemini(wd, config, system, user, endpoint.baseUrl, {
          omitTemperature: endpoint.omitTemperature,
          headers: endpoint.headers,
        })
      case 'openai-compatible':
        return chatOpenAiCompatible(wd, endpoint.baseUrl, config, system, user, {
          omitTemperature: endpoint.omitTemperature,
          bodyExtras: endpoint.bodyExtras,
          headers: endpoint.headers,
        })
      case 'openai-responses':
        return chatOpenAiResponses(wd, endpoint.baseUrl, config, system, user, {
          bodyExtras: endpoint.bodyExtras,
          headers: endpoint.headers,
        })
    }
  })
  return result.catch((e) => {
    if (e instanceof ResponseBodyTooLargeError) return { ok: false as const, error: e.message }
    throw e
  })
}
