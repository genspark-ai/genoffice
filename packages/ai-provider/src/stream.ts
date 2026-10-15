import type { AgentMessage, AgentToolDef } from '@genoffice/agent-core'
import { withOutputCapFallback } from './output-cap'
import { streamAnthropic } from './protocols/anthropic'
import { streamGemini } from './protocols/gemini'
import { streamOpenAiCompatible } from './protocols/openai-compatible'
import { streamCodexAppServer } from './codex-app-server'
import type { StreamCallbacks } from './protocols/shared'
import { getProviderAdapter, type AiProtocol } from './registry'
import type { AiProviderConfig, AiProviderId } from './types'

export { streamAnthropic } from './protocols/anthropic'
export { streamGemini } from './protocols/gemini'
export { streamOpenAiCompatible } from './protocols/openai-compatible'
export { AiCreditsError, sseLines } from './protocols/shared'
export type { StreamCallbacks } from './protocols/shared'

/** route a streaming, tool-calling-capable turn by provider id */
export async function streamForProvider(
  provider: AiProviderId,
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
): Promise<void> {
  const endpoint = getProviderAdapter(provider).resolveEndpoint(config)
  const { baseUrl } = endpoint
  if (endpoint.model) config = { ...config, model: endpoint.model }
  if (endpoint.protocol === 'codex-app-server') {
    return streamCodexAppServer(config, system, messages, tools, maxTokens, cb)
  }
  // A model the user declared a ceiling for outranks the caller's budget; the
  // endpoint's own 400 would otherwise cost a billed round trip.
  const cap = endpoint.maxTokens !== undefined ? Math.min(maxTokens, endpoint.maxTokens) : maxTokens
  const protocol: Exclude<AiProtocol, 'codex-app-server'> = endpoint.protocol
  return withOutputCapFallback(baseUrl, config.model, cap, (tokens) => {
    switch (protocol) {
      case 'anthropic':
        return streamAnthropic(config, system, messages, tools, tokens, cb, baseUrl, {
          headers: endpoint.headers,
          idleTimeoutMs: endpoint.idleTimeoutMs,
        })
      case 'gemini':
        return streamGemini(config, system, messages, tools, tokens, cb, baseUrl, {
          omitTemperature: endpoint.omitTemperature,
          headers: endpoint.headers,
          idleTimeoutMs: endpoint.idleTimeoutMs,
        })
      case 'openai-compatible':
        return streamOpenAiCompatible(baseUrl, config, system, messages, tools, tokens, cb, {
          omitTemperature: endpoint.omitTemperature,
          useMaxCompletionTokens: endpoint.useMaxCompletionTokens,
          bodyExtras: endpoint.bodyExtras,
          echoReasoning: endpoint.echoReasoning,
          headers: endpoint.headers,
          idleTimeoutMs: endpoint.idleTimeoutMs,
        })
    }
  })
}
