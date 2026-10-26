// Bridge between GenOffice's harness contracts and pi-agent-core/pi-ai.
// Kept in one package so the cutover (or rollback) is a single import site.
import {
  createModels,
  createProvider,
  type Models,
  type MutableModels,
  type Model,
  type Api,
} from '@earendil-works/pi-ai'
import { anthropicMessagesApi } from '@earendil-works/pi-ai/api/anthropic-messages.lazy'
import { googleGenerativeAIApi } from '@earendil-works/pi-ai/api/google-generative-ai.lazy'
import { openAICompletionsApi } from '@earendil-works/pi-ai/api/openai-completions.lazy'
import { openAIResponsesApi } from '@earendil-works/pi-ai/api/openai-responses.lazy'
import type {
  AssistantMessageEventStream,
  Context,
  Tool as PiTool,
  ToolCall,
  Message as PiMessage,
  TextContent,
  ImageContent,
  ToolResultMessage,
} from '@earendil-works/pi-ai'
import type { AiTokenUsage } from '@genoffice/ai-provider'
import type {
  AgentMessage,
  AgentToolDef,
} from '@genoffice/agent-core'
import type {
  AiSettings,
  StreamCallbacks,
} from '@genoffice/ai-provider'
import {
  activeProvider,
  getProviderAdapter,
  clampMaxOutputTokens,
} from '@genoffice/ai-provider'
import { endpointSupportsVision } from '@genoffice/ai-provider/browser'

// ---------------------------------------------------------------------------
// AiSettings → pi Model
// ---------------------------------------------------------------------------

let cachedModels: Models | undefined

function modelsInstance(): Models {
  if (!cachedModels) cachedModels = createModels()
  return cachedModels
}

/** Map GenOffice provider config onto a pi Model for one request. */
export function piModelFor(settings: AiSettings): Model<Api> {
  const providerId = activeProvider(settings)
  const adapter = getProviderAdapter(providerId)
  const config = settings.providers[providerId]
  const endpoint = adapter.resolveEndpoint(config)

  // GenOffice custom endpoints and most providers talk OpenAI-compatible.
  // Anthropic/Gemini have dedicated pi protocols.
  let api: Api
  switch (providerId) {
    case 'anthropic':
      api = 'anthropic-messages'
      break
    case 'gemini':
      api = 'google-generative-ai'
      break
    case 'codex':
      api = 'openai-codex-responses'
      break
    default:
      api = 'openai-completions'
  }

  const model: Model<Api> = {
    id: endpoint.model ?? config.model ?? '',
    name: endpoint.model ?? config.model ?? '',
    api,
    provider: providerId,
    baseUrl: endpoint.baseUrl,
    input: endpointSupportsVision(settings) ? ['text', 'image'] : ['text'],
    cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
  } as Model<Api>

  // Register a minimal static-auth provider so Models.streamSimple can resolve
  // credentials for any GenOffice provider id (custom endpoints included).
  const models = modelsInstance() as MutableModels
  if (!models.getProvider(providerId)) {
    const streams =
      api === 'anthropic-messages'
        ? anthropicMessagesApi()
        : api === 'google-generative-ai'
          ? googleGenerativeAIApi()
          : api === 'openai-codex-responses'
            ? openAIResponsesApi()
            : openAICompletionsApi()
    models.setProvider(
      createProvider({
        id: providerId,
        baseUrl: endpoint.baseUrl,
        models: [model],
        api: streams,
        auth: {
          apiKey: {
            name: `${providerId} API key`,
            resolve: async () =>
              config.apiKey
                ? { auth: { apiKey: config.apiKey } }
                : undefined,
          },
        },
      }),
    )
  }

  return model
}

// ---------------------------------------------------------------------------
// AgentToolDef (plain JSON Schema) → pi Tool
// ---------------------------------------------------------------------------

export function piTools(tools: AgentToolDef[]): PiTool[] {
  return tools.map((t) => ({
    name: t.name,
    description: t.description,
    // pi serializes parameters as JSON Schema verbatim on the wire; TSchema is
    // a compile-time brand over plain JSON Schema, so a cast suffices.
    parameters: t.inputSchema as PiTool['parameters'],
  }))
}

// ---------------------------------------------------------------------------
// AgentMessage[] → pi Message[]
// ---------------------------------------------------------------------------

export function piMessages(messages: AgentMessage[]): PiMessage[] {
  const out: PiMessage[] = []
  for (const m of messages) {
    if (m.role === 'user') {
      const content: (TextContent | ImageContent)[] = []
      if (m.text) content.push({ type: 'text', text: m.text })
      for (const img of m.images ?? []) {
        content.push({ type: 'image', data: img.base64, mimeType: img.mime })
      }
      out.push({ role: 'user', content } as PiMessage)
    } else if (m.role === 'assistant') {
      const content: (TextContent | ToolCall)[] = []
      if (m.text) content.push({ type: 'text', text: m.text })
      for (const tc of m.toolCalls ?? []) {
        content.push({
          type: 'toolCall',
          id: tc.id,
          name: tc.name,
          arguments: (tc.input ?? {}) as ToolCall['arguments'],
        })
      }
      if (content.length) out.push({ role: 'assistant', content } as PiMessage)
    } else if (m.role === 'tool') {
      // GenOffice batches tool results per turn; pi expects one message per call.
      for (const r of m.results) {
        out.push({
          role: 'toolResult',
          toolCallId: r.id,
          toolName: r.name,
          content: [{ type: 'text', text: r.output }],
          isError: r.isError ?? false,
        } as ToolResultMessage)
      }
    }
  }
  return out
}

// ---------------------------------------------------------------------------
// One-shot chat (ai:chat channel), backed by pi-ai streamSimple collected to
// completion. Same contract as chatForProvider: failures report as { ok:false }.
// ---------------------------------------------------------------------------

export async function chatViaPi(
  settings: AiSettings,
  system: string,
  user: string,
): Promise<{ ok: boolean; content?: string; error?: string }> {
  try {
    const model = piModelFor(settings)
    const stream = modelsInstance().streamSimple(model, {
      systemPrompt: system,
      messages: [{ role: 'user', content: [{ type: 'text', text: user }] } as PiMessage],
    })
    let text = ''
    for await (const ev of stream) {
      if (ev.type === 'text_delta') text += ev.delta
      else if (ev.type === 'error') throw new Error(ev.error.errorMessage ?? 'stream error')
    }
    return { ok: true, content: text }
  } catch (err) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) }
  }
}

// ---------------------------------------------------------------------------
// Drop-in replacement for streamForProvider, backed by pi-ai.
// ---------------------------------------------------------------------------

export async function streamViaPi(
  settings: AiSettings,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
): Promise<void> {
  const model = piModelFor(settings)
  const context: Context = {
    systemPrompt: system,
    messages: piMessages(messages),
    tools: piTools(tools),
  }
  const stream: AssistantMessageEventStream = modelsInstance().streamSimple(model, context, {
    maxTokens: clampMaxOutputTokens(maxTokens),
    signal: cb.signal,
  })

  try {
    for await (const ev of stream) {
      switch (ev.type) {
        case 'text_delta':
          cb.onDelta(ev.delta)
          break
        case 'thinking_delta':
          cb.onReasoningDelta?.(ev.delta)
          break
        case 'toolcall_end': {
          // toolcall_end carries the complete ToolCall; earlier deltas are not
          // needed because GenOffice's transport only wants the final call.
          const tc = ev.toolCall
          cb.onToolCall({
            id: tc.id,
            name: tc.name,
            input: tc.arguments as Record<string, unknown>,
          })
          break
        }
        case 'done': {
          if (ev.message.usage) {
            const u = ev.message.usage
            cb.onUsage?.({ inputTokens: u.input, outputTokens: u.output })
          }
          cb.onStopReason?.(ev.reason)
          break
        }
        case 'error': {
          throw new Error(ev.error.errorMessage ?? 'stream error')
        }
      }
    }
  } catch (err) {
    // streamForProvider throws on transport failure; the IPC layer maps it to
    // an error chunk. Keep the same contract here.
    throw err instanceof Error ? err : new Error(String(err))
  }
}
