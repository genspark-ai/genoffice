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
import type { AgentMessage, AgentToolDef } from '@genoffice/agent-core'
import type { AiSettings, StreamCallbacks } from '@genoffice/ai-provider'
import {
  AI_CHAT_RESPONSE_TIMEOUT_MS,
  activeProvider,
  getProviderAdapter,
  clampMaxOutputTokens,
  AiCreditsError,
  createStreamWatchdog,
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
  // Protocol gap frozen by the migration charter (principle 4): GenOffice's
  // codex provider speaks the local codex-app-server subprocess protocol; pi
  // only offers the HTTP Responses API, which is not isomorphic. Fail loudly
  // instead of attempting an HTTP call that can only end in a confusing
  // auth/URL error. Rollback for codex users: GENOFFICE_AI_BACKEND=legacy.
  if (providerId === 'codex') {
    throw new Error(
      'The codex provider is not supported on the pi backend; set GENOFFICE_AI_BACKEND=legacy to use it.',
    )
  }
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
            resolve: async () => (config.apiKey ? { auth: { apiKey: config.apiKey } } : undefined),
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
// pi failure events → GenOffice typed errors
// ---------------------------------------------------------------------------

// pi ends every failed stream with an `error` event whose payload is an
// AssistantMessage carrying only the composed `errorMessage` string — the
// underlying SDK error object (HTTP status field, abort name, cause chain)
// never crosses the stream boundary. The helpers below read that text the same
// way the legacy protocol streams read their error bodies, so the apps'
// errorCode classification (timeout / credits / network / overloaded / region)
// behaves identically behind either backend. Watchdog timeouts are not mapped
// here: wd.guard rewrites those into AiTimeoutError before this code runs.

/**
 * pi's formatProviderError puts a leading HTTP status on provider failures in
 * one of three shapes: "NNN: {body}" when the SDK message lacks the body,
 * "NNN {body}" when it already carries it, and `NNN "plain body"` when the
 * body is not JSON. All of them normalize to the legacy "HTTP NNN: …" form.
 */
const PI_STATUS_PREFIX = /^\s*(\d{3})(?:\s*:\s*|\s+)/

/**
 * Credits-exhaustion notices, mirroring the legacy openai-completions
 * creditsNoticeText wording (order-insensitive so both "credits … exhausted"
 * and "insufficient … credits" match, plus the Genspark pricing pointer).
 */
const PI_CREDITS_NOTICE =
  /(genspark\.ai\/pricing)|(credits?[^\n]{0,120}(exhausted|insufficient))|((exhausted|insufficient)[^\n]{0,120}credits?)/i

/**
 * Map one pi `error` event onto the error the legacy stack would have thrown
 * for the same failure, so the callers' instanceof/predicate classification
 * (docs-main.ts timeout/credits/network/overloaded/region) needs no
 * pi-specific branches:
 * - a credits notice, or an HTTP 402 (billing) → AiCreditsError ('credits')
 * - an HTTP status prefix is rewritten to the legacy "HTTP NNN: …" form, which
 *   is the shape isAiOverloadedError / isAiRegionBlockedError / the network
 *   classifier pattern-match on in the legacy streams
 * - anything else (including aborts — the IPC layer recognizes those via the
 *   request signal before it ever looks at error classes) stays a plain Error
 *   carrying pi's message
 */
export function mapPiErrorEvent(errorMessage?: string): Error {
  const raw = errorMessage?.trim() || 'stream error'
  const status = PI_STATUS_PREFIX.exec(raw)
  if (status?.[1] === '402' || PI_CREDITS_NOTICE.test(raw)) return new AiCreditsError(raw)
  return new Error(status ? `HTTP ${status[1]}: ${raw.slice(status[0].length)}` : raw)
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
    // Same budget chatForProvider gives a one-shot chat: a gateway may answer a
    // streaming request with one buffered body (then the whole generation
    // elapses before the first event), so the connect phase gets the long
    // AI_CHAT_RESPONSE_TIMEOUT_MS and the idle phase the watchdog default.
    const wd = createStreamWatchdog(undefined, AI_CHAT_RESPONSE_TIMEOUT_MS)
    let text = ''
    await wd.guard(async () => {
      const stream = modelsInstance().streamSimple(
        model,
        {
          systemPrompt: system,
          messages: [{ role: 'user', content: [{ type: 'text', text: user }] } as PiMessage],
        },
        { signal: wd.signal },
      )
      for await (const ev of stream) {
        wd.touch()
        if (ev.type === 'text_delta') text += ev.delta
        else if (ev.type === 'error') throw mapPiErrorEvent(ev.error.errorMessage)
      }
    })
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
  // pi's own lifecycle knobs do not match GenOffice's watchdog budgets: pi maps
  // `timeoutMs` to the provider SDK's request timeout (OpenAI/Anthropic SDKs:
  // time to response headers, 10-minute default, cleared once they arrive) and
  // has no idle timeout of its own — post-header idleness is left to Node's
  // fetch defaults (undici bodyTimeout, 5 minutes). Neither equals the legacy
  // 60s connect / 180s idle budget, and arming timeoutMs at 180s would cap a
  // turn's total wall time, re-creating the false kills that moved the legacy
  // watchdog to a re-armed idle budget in the first place. So keep the legacy
  // watchdog and align through pi's `signal` option instead: wd.signal aborts
  // pi's in-flight request, wd.guard surfaces the trip as AiTimeoutError.
  const wd = createStreamWatchdog(cb.signal)

  try {
    await wd.guard(async () => {
      const stream: AssistantMessageEventStream = modelsInstance().streamSimple(model, context, {
        maxTokens: clampMaxOutputTokens(maxTokens),
        signal: wd.signal,
      })

      for await (const ev of stream) {
        // every event is parsed provider output — bytes were on the wire
        wd.touch()
        switch (ev.type) {
          // Events with no callback of their own still count as activity for
          // the host: `start` is the headers-arrived ping (legacy touches the
          // watchdog and pings the renderer at that exact moment), and
          // toolcall deltas are the only signal during long tool-argument
          // generation, which otherwise stays silent until toolcall_end.
          case 'start':
          case 'toolcall_start':
          case 'toolcall_delta':
            cb.onActivity?.()
            break
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
            throw mapPiErrorEvent(ev.error.errorMessage)
          }
        }
      }
    })
  } catch (err) {
    // streamForProvider throws on transport failure; the IPC layer maps it to
    // an error chunk. Keep the same contract here.
    throw err instanceof Error ? err : new Error(String(err))
  }
}
