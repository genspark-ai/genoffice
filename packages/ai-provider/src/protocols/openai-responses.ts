import type { AgentMessage, AgentToolDef } from '@genoffice/agent-core'
import { aiFetch } from '../fetch'
import { httpBodyDetail } from '../http-error'
import { gensparkAttributionHeaders, opencodeSessionHeaders } from '../providers'
import type { AiChatResponse, AiProviderConfig, AiTokenUsage } from '../types'
import { createStreamWatchdog, type StreamWatchdog } from '../watchdog'
import {
  endpointUrl,
  jsonBodyInsteadOfSse,
  parseToolInput,
  readCappedResponseText,
  sseDataEvents,
  throwIfCreditsNotice,
  throwIfToolJsonOverBudget,
  usageCount,
  type StreamCallbacks,
} from './shared'

/**
 * OpenAI Responses API wire protocol (`POST /v1/responses`). The Chat
 * Completions endpoint serves the GPT-5.6 family's function calling only with
 * `reasoning_effort: 'none'` (their 400 names both this API and effort none),
 * and the GPT-6 family needs this API for tool calling at all — so models on
 * the reasoning line route here, keeping their default reasoning beside tools.
 *
 * Shape differences from chat/completions that this module owns: prior turns
 * replay as typed input items (function_call / function_call_output pairs
 * matched by `call_id`), tools are flat `{type:'function', name, …}` objects,
 * the system prompt travels as `instructions`, output streams as typed events
 * (`response.output_text.delta`, `response.function_call_arguments.*`,
 * `response.completed` carrying usage), and there is no finish_reason — an
 * output-capped turn ends as `response.incomplete`.
 */

/** the reasoning-line models Chat Completions serves only with effort none */
export const OPENAI_RESPONSES_MODEL_RE = /(^|\/)gpt-5\.6([-./]|$)/i

export interface OpenAiResponsesRequestOptions {
  headers?: Record<string, string> | undefined
  idleTimeoutMs?: number | undefined
  /** pin the reasoning effort; omitted keeps the model's default (recommended) */
  reasoningEffort?: string | undefined
}

function isFunctionCallItem(item: unknown): item is {
  type: 'function_call'
  id?: string
  call_id?: string
  name?: string
  arguments?: string
} {
  return (
    typeof item === 'object' &&
    item !== null &&
    (item as { type?: unknown }).type === 'function_call'
  )
}

/** chat history → typed input items (system travels as `instructions`) */
export function responsesInputItems(messages: AgentMessage[]): unknown[] {
  const out: unknown[] = []
  for (const m of messages) {
    if (m.role === 'user') {
      if (!m.images?.length) {
        out.push({ role: 'user', content: m.text })
      } else {
        out.push({
          role: 'user',
          content: [
            ...(m.text ? [{ type: 'input_text', text: m.text }] : []),
            ...m.images.map((img) => ({
              type: 'input_image',
              image_url: `data:${img.mime};base64,${img.base64}`,
            })),
          ],
        })
      }
    } else if (m.role === 'assistant') {
      // an assistant turn replays as its output items: text first, then one
      // function_call per call — the items the model itself would have emitted
      if (m.text) {
        out.push({
          role: 'assistant',
          content: [{ type: 'output_text', text: m.text }],
        })
      }
      for (const call of m.toolCalls ?? []) {
        out.push({
          type: 'function_call',
          call_id: call.id,
          name: call.name,
          arguments: JSON.stringify(call.input),
        })
      }
    } else {
      for (const r of m.results) {
        // OpenAI has no structured is_error flag (Anthropic sends one); prefix
        // the content so the model can tell a failed call from a success.
        out.push({
          type: 'function_call_output',
          call_id: r.id,
          output: r.isError ? `Error: ${r.output}` : r.output,
        })
      }
    }
  }
  return out
}

function emitCompletedFunctionCalls(output: unknown, cb: StreamCallbacks): void {
  if (!Array.isArray(output)) return
  for (const item of output) {
    if (!isFunctionCallItem(item) || !item.call_id || !item.name) continue
    const { input, error } = parseToolInput(item.arguments ?? '')
    cb.onToolCall({
      id: item.call_id,
      name: item.name,
      input,
      inputError: error,
    })
  }
}

function usageOf(response: unknown): AiTokenUsage | undefined {
  const usage = (response as { usage?: Record<string, unknown> } | undefined)?.usage
  if (!usage || typeof usage !== 'object') return undefined
  const input = usageCount((usage as { input_tokens?: unknown }).input_tokens)
  const output = usageCount((usage as { output_tokens?: unknown }).output_tokens)
  return {
    ...(input !== undefined ? { inputTokens: input } : {}),
    ...(output !== undefined ? { outputTokens: output } : {}),
  }
}

/** the terminal `response.completed` / `response.incomplete` payloads */
function emitTerminal(
  event: { type: string; response?: Record<string, unknown> },
  cb: StreamCallbacks,
): void {
  const response = event.response
  if (!response) return
  emitCompletedFunctionCalls(response.output, cb)
  if (event.type === 'response.incomplete') {
    // no finish_reason on this API: the output cap ends the turn here
    cb.onStopReason?.('max_tokens')
  }
}

async function responsesTurn(
  baseUrl: string,
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
  wd: StreamWatchdog,
  options: OpenAiResponsesRequestOptions,
): Promise<void> {
  const onBytes = () => {
    wd.touch()
    cb.onActivity?.()
  }
  const response = await aiFetch(endpointUrl(baseUrl, 'responses'), {
    method: 'POST',
    signal: wd.signal,
    headers: {
      'Content-Type': 'application/json',
      ...(config.apiKey ? { Authorization: `Bearer ${config.apiKey}` } : {}),
      ...gensparkAttributionHeaders(baseUrl),
      ...opencodeSessionHeaders(baseUrl, cb.sessionId),
      ...options.headers,
    },
    body: JSON.stringify({
      model: config.model,
      instructions: system,
      input: responsesInputItems(messages),
      ...(tools.length > 0
        ? {
            // flat function tools; strict:false — agent schemas are not
            // guaranteed to satisfy the strict-mode subset
            tools: tools.map((t) => ({
              type: 'function',
              name: t.name,
              description: t.description,
              parameters: t.inputSchema,
              strict: false,
            })),
          }
        : {}),
      // effort stays the model's default: the point of this protocol is to
      // keep reasoning beside tools instead of pinning it away
      ...(options.reasoningEffort ? { reasoning: { effort: options.reasoningEffort } } : {}),
      max_output_tokens: maxTokens,
      store: false,
      stream: true,
    }),
  })
  onBytes()
  if (!response.ok || !response.body) {
    throw new Error(
      `HTTP ${response.status}: ${httpBodyDetail(await readCappedResponseText(response, onBytes))}`,
    )
  }
  // a gateway that swallowed the stream answers one whole response object
  const jsonBody = await jsonBodyInsteadOfSse(response, onBytes)
  if (jsonBody !== null) {
    return emitWholeResponse(jsonBody, cb)
  }

  const pending = new Map<string, { callId: string; name: string; args: string }>()
  let emittedText = false
  let emittedTool = false
  let turnUsage: AiTokenUsage | undefined
  let terminated = false

  const finalizeCall = (raw: {
    call_id?: string
    name?: string
    arguments?: string
    id?: string
  }): boolean => {
    const callId = raw.call_id ?? ''
    const name = raw.name ?? ''
    if (!callId || !name) return false
    throwIfToolJsonOverBudget((raw.arguments ?? '').length, 'openai-responses')
    const { input, error } = parseToolInput(raw.arguments ?? '')
    emittedTool = true
    cb.onToolCall({ id: callId, name, input, inputError: error })
    return true
  }

  for await (const sse of sseDataEvents(response.body, onBytes)) {
    if (sse.raw === '[DONE]') break
    if (sse.json === undefined) continue
    const event = sse.json as { type?: string } & Record<string, unknown>
    const type = event.type ?? ''
    switch (type) {
      case 'response.output_item.added': {
        const item = event.item as Record<string, unknown> | undefined
        if (isFunctionCallItem(item)) {
          const key = String(item.id ?? item.call_id ?? '')
          if (key) {
            pending.set(key, {
              callId: String(item.call_id ?? ''),
              name: String(item.name ?? ''),
              args: String(item.arguments ?? ''),
            })
          }
        }
        break
      }
      case 'response.function_call_arguments.delta': {
        const entry = pending.get(String(event.item_id ?? ''))
        if (entry) entry.args += String(event.delta ?? '')
        break
      }
      case 'response.output_item.done': {
        const item = event.item as Record<string, unknown> | undefined
        if (isFunctionCallItem(item)) {
          // the done item is complete; the streamed fragment is redundant
          finalizeCall(item)
          const key = String(item.id ?? item.call_id ?? '')
          pending.delete(key)
        }
        break
      }
      case 'response.output_text.delta': {
        const text = String(event.delta ?? '')
        if (text) {
          emittedText = true
          cb.onDelta(text)
        }
        break
      }
      case 'response.reasoning_summary_text.delta':
      case 'response.reasoning_text.delta': {
        const text = String(event.delta ?? '')
        if (text) cb.onReasoningDelta?.(text)
        break
      }
      case 'response.completed':
      case 'response.incomplete': {
        emitTerminal(event as never, cb)
        turnUsage = usageOf(event.response as Record<string, unknown>)
        terminated = true
        break
      }
      case 'response.failed': {
        const response = event.response as { error?: { message?: string } } | undefined
        throw new Error(response?.error?.message || 'The Responses API reported a failed turn')
      }
      case 'error': {
        const payload = event as { message?: string; code?: string }
        throw new Error(payload.message || payload.code || 'The Responses API reported an error')
      }
      default:
        break
    }
  }
  // a pending call whose done event never arrived: emit what streamed
  for (const entry of pending.values()) {
    finalizeCall({ call_id: entry.callId, name: entry.name, arguments: entry.args })
  }
  if (!terminated) {
    // the stream ended without a terminal event (truncated connection): the
    // watchdog owns idle hangs, this is a body-level truncation
    throw new Error('The Responses API stream ended without a terminal event')
  }
  if (!emittedText && !emittedTool) {
    throw new Error('The model returned no content (empty stream)')
  }
  if (turnUsage) cb.onUsage?.(turnUsage)
}

/** a whole (non-streamed) response object delivered as one JSON body */
function emitWholeResponse(bodyText: string, cb: StreamCallbacks): void {
  let json: { output?: unknown; status?: string; usage?: unknown; error?: { message?: string } }
  try {
    json = JSON.parse(bodyText)
  } catch {
    throw new Error(`AI returned a non-JSON response: ${httpBodyDetail(bodyText)}`)
  }
  throwIfCreditsNotice(bodyText)
  if (json.error) throw new Error(json.error.message || 'The Responses API reported an error')
  emitCompletedFunctionCalls(json.output, cb)
  const text = (Array.isArray(json.output) ? json.output : [])
    .flatMap((item) =>
      typeof item === 'object' && item !== null && (item as { type?: unknown }).type === 'message'
        ? ((item as { content?: unknown }).content ?? [])
        : [],
    )
    .flatMap((part) =>
      typeof part === 'object' &&
      part !== null &&
      (part as { type?: unknown }).type === 'output_text'
        ? String((part as { text?: unknown }).text ?? '')
        : [],
    )
    .join('')
  if (text) cb.onDelta(text)
  const usage = usageOf(json as Record<string, unknown>)
  if (json.status === 'incomplete') cb.onStopReason?.('max_tokens')
  if (usage) cb.onUsage?.(usage)
  const hasCall = Array.isArray(json.output) && json.output.some(isFunctionCallItem)
  if (!text && !hasCall) {
    throw new Error('The model returned no content (empty response)')
  }
}

export async function streamOpenAiResponses(
  baseUrl: string,
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
  options: OpenAiResponsesRequestOptions = {},
): Promise<void> {
  const wd = createStreamWatchdog(cb.signal, undefined, options.idleTimeoutMs)
  return wd.guard(() =>
    responsesTurn(baseUrl, config, system, messages, tools, maxTokens, cb, wd, options),
  )
}

/** non-streaming single question over the Responses API (connectivity probes) */
export async function chatOpenAiResponses(
  wd: StreamWatchdog,
  baseUrl: string,
  config: AiProviderConfig,
  system: string,
  user: string,
  options: OpenAiResponsesRequestOptions = {},
): Promise<AiChatResponse> {
  const response = await aiFetch(endpointUrl(baseUrl, 'responses'), {
    method: 'POST',
    signal: wd.signal,
    headers: {
      'Content-Type': 'application/json',
      ...(config.apiKey ? { Authorization: `Bearer ${config.apiKey}` } : {}),
      ...gensparkAttributionHeaders(baseUrl),
      ...opencodeSessionHeaders(baseUrl),
      ...options.headers,
    },
    body: JSON.stringify({
      model: config.model,
      instructions: system,
      input: user,
      max_output_tokens: 4096,
      store: false,
    }),
  })
  wd.touch()
  if (!response.ok) {
    return {
      ok: false,
      error: `HTTP ${response.status}: ${httpBodyDetail(await readCappedResponseText(response, () => wd.touch()))}`,
    }
  }
  const bodyText = await readCappedResponseText(response, () => wd.touch())
  let json: {
    output?: Array<{ type?: string; content?: Array<{ type?: string; text?: string }> }>
    error?: { message?: string }
  }
  try {
    json = JSON.parse(bodyText)
  } catch {
    return { ok: false, error: `AI returned a non-JSON response: ${httpBodyDetail(bodyText)}` }
  }
  if (json.error)
    return { ok: false, error: json.error.message ?? 'The Responses API reported an error' }
  const content = (json.output ?? [])
    .filter((item) => item.type === 'message')
    .flatMap((item) => item.content ?? [])
    .filter((part) => part.type === 'output_text')
    .map((part) => part.text ?? '')
    .join('')
  if (!content) return { ok: false, error: 'AI returned an empty response' }
  return { ok: true, content }
}
