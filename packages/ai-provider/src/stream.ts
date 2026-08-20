import type { AgentMessage, AgentToolDef } from '@genoffice/agent-core'
import { streamAnthropic } from './protocols/anthropic'
import { streamGemini } from './protocols/gemini'
import { streamOpenAiCompatible } from './protocols/openai-compatible'
import type { StreamCallbacks } from './protocols/shared'
import { getProviderAdapter } from './registry'
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
  switch (endpoint.protocol) {
    case 'anthropic':
      return streamAnthropic(config, system, messages, tools, maxTokens, cb, baseUrl)
    case 'gemini':
      return streamGemini(config, system, messages, tools, maxTokens, cb, baseUrl)
    case 'openai-compatible':
      return streamOpenAiCompatible(
        baseUrl,
        config,
        system,
        messages,
        tools,
        maxTokens,
        cb,
        endpoint.omitTemperature,
      )
  }
}
