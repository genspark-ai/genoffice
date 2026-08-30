import type { AgentMessage, AgentToolCall, AgentToolDef } from '@genoffice/agent-core'
import { aiFetch } from '../fetch'
import { httpBodyDetail } from '../http-error'
import { gensparkAttributionHeaders } from '../providers'
import { modelEchoesReasoning } from '../registry'
import type { AiChatResponse, AiProviderConfig } from '../types'
import { createStreamWatchdog, type StreamWatchdog } from '../watchdog'
import {
  jsonBodyInsteadOfSse,
  parseToolInput,
  sseErrorText,
  sseLines,
  throwIfCreditsNotice,
  type StreamCallbacks,
} from './shared'

function openAiMessages(
  system: string,
  messages: AgentMessage[],
  echoReasoning: boolean,
): unknown[] {
  const out: unknown[] = [{ role: 'system', content: system }]
  for (const m of messages) {
    if (m.role === 'user') {
      if (!m.images?.length) {
        out.push({ role: 'user', content: m.text })
      } else {
        out.push({
          role: 'user',
          content: [
            ...(m.text ? [{ type: 'text', text: m.text }] : []),
            ...m.images.map((img) => ({
              type: 'image_url',
              image_url: { url: `data:${img.mime};base64,${img.base64}` },
            })),
          ],
        })
      }
    } else if (m.role === 'assistant') {
      const hasTools = !!(m.toolCalls && m.toolCalls.length > 0)
      // content:null with no tool_calls is an empty assistant turn; some OpenAI-
      // compatible proxies drop or reject the follow-up conversation after that.
      out.push({
        role: 'assistant',
        content: m.text || (hasTools ? null : '(no content)'),
        ...(echoReasoning && m.reasoning ? { reasoning_content: m.reasoning } : {}),
        ...(hasTools
          ? {
              tool_calls: m.toolCalls!.map((call) => ({
                id: call.id,
                type: 'function',
                function: { name: call.name, arguments: JSON.stringify(call.input) },
              })),
            }
          : {}),
      })
    } else {
      for (const r of m.results) {
        out.push({ role: 'tool', tool_call_id: r.id, content: r.output })
      }
    }
  }
  return out
}

/** Emits a complete (non-streamed) chat completion delivered as a plain JSON body. */
function emitOpenAiJsonMessage(bodyText: string, cb: StreamCallbacks): void {
  let msg: {
    choices?: Array<{
      message?: {
        content?: string | null
        reasoning_content?: string
        tool_calls?: Array<{ id?: string; function?: { name?: string; arguments?: string } }>
      }
      finish_reason?: string | null
    }>
    error?: { message?: string } | string
  }
  try {
    msg = JSON.parse(bodyText) as typeof msg
  } catch {
    throw new Error(`The model returned an unparseable JSON body: ${httpBodyDetail(bodyText)}`)
  }
  if (msg.error) throw new Error(sseErrorText(msg.error, 'Model error'))
  const choice = msg.choices?.[0]
  let emitted = false
  if (choice?.message?.reasoning_content) cb.onReasoningDelta?.(choice.message.reasoning_content)
  if (choice?.message?.content) {
    emitted = true
    cb.onDelta(choice.message.content)
  }
  const toolCalls: AgentToolCall[] = []
  for (const tc of choice?.message?.tool_calls ?? []) {
    if (!tc.function?.name) continue
    emitted = true
    const { input, error } = parseToolInput(tc.function.arguments ?? '')
    toolCalls.push({
      id: tc.id ?? crypto.randomUUID(),
      name: tc.function.name,
      input,
      inputError: error,
    })
  }
  // a 'length' finish may have cut off the last tool call's arguments
  const lastTool = toolCalls.at(-1)
  if (choice?.finish_reason === 'length' && lastTool) lastTool.truncated = true
  for (const call of toolCalls) cb.onToolCall(call)
  if (!emitted) throw new Error(`The model returned no content: ${httpBodyDetail(bodyText)}`)
  if (choice?.finish_reason === 'length') cb.onStopReason?.('max_tokens')
}

/** Per-endpoint request shaping resolved from the provider registry. */
export interface OpenAiRequestOptions {
  omitTemperature?: boolean | undefined
  /** send the output cap as OpenAI's renamed `max_completion_tokens` (GPT-5.x/o-series 400 on `max_tokens`) */
  useMaxCompletionTokens?: boolean | undefined
  /** vendor-specific fields merged into the request body (e.g. DeepSeek's `thinking`) */
  bodyExtras?: Record<string, unknown> | undefined
}

export async function streamOpenAiCompatible(
  baseUrl: string,
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
  options: OpenAiRequestOptions = {},
): Promise<void> {
  const wd = createStreamWatchdog(cb.signal)
  return wd.guard(() =>
    openAiCompatibleTurn(baseUrl, config, system, messages, tools, maxTokens, cb, wd, options),
  )
}

async function openAiCompatibleTurn(
  baseUrl: string,
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
  wd: StreamWatchdog,
  options: OpenAiRequestOptions,
): Promise<void> {
  const onBytes = () => {
    wd.touch()
    cb.onActivity?.()
  }
  const response = await aiFetch(`${baseUrl.replace(/\/$/, '')}/chat/completions`, {
    method: 'POST',
    signal: wd.signal,
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${config.apiKey}`,
      ...gensparkAttributionHeaders(baseUrl),
    },
    body: JSON.stringify({
      model: config.model,
      ...(options.useMaxCompletionTokens
        ? { max_completion_tokens: maxTokens }
        : { max_tokens: maxTokens }),
      messages: openAiMessages(system, messages, modelEchoesReasoning(config.model)),
      ...(tools.length > 0
        ? {
            tools: tools.map((t) => ({
              type: 'function',
              function: { name: t.name, description: t.description, parameters: t.inputSchema },
            })),
          }
        : {}),
      ...(options.omitTemperature ? {} : { temperature: 0.3 }),
      ...options.bodyExtras,
      stream: true,
    }),
  })
  // headers arrived: ping the renderer watchdog too, or a slow first chunk could trip it
  onBytes()
  if (!response.ok || !response.body) {
    throw new Error(`HTTP ${response.status}: ${httpBodyDetail(await response.text())}`)
  }
  const jsonBody = await jsonBodyInsteadOfSse(response)
  if (jsonBody !== null) {
    throwIfCreditsNotice(jsonBody)
    return emitOpenAiJsonMessage(jsonBody, cb)
  }
  // tool call arguments stream in fragments keyed by index
  const pendingTools = new Map<number, { id: string; name: string; json: string }>()
  let stopReason: string | undefined
  let abnormalFinish: string | undefined
  let sawFinish = false
  let emitted = false
  const flushTools = () => {
    const entries = [...pendingTools.entries()].sort(([a], [b]) => a - b)
    const lastIndex = entries.at(-1)?.[0]
    for (const [index, pending] of entries) {
      if (pending.name) {
        const { input, error } = parseToolInput(pending.json)
        emitted = true
        cb.onToolCall({
          id: pending.id,
          name: pending.name,
          input,
          inputError: error,
          // a 'length' finish cuts off the last streaming tool's arguments
          ...(stopReason === 'max_tokens' && index === lastIndex ? { truncated: true } : {}),
        })
      }
    }
    pendingTools.clear()
  }
  for await (const line of sseLines(response.body, onBytes)) {
    if (!line.startsWith('data:')) continue
    const payload = line.slice(5).trim()
    if (!payload) continue
    if (payload === '[DONE]') break
    // A truncated frame or a non-JSON keep-alive from a proxy should skip
    // that event, not kill the entire AI turn with a parser error.
    let event
    try {
      event = JSON.parse(payload) as {
        choices?: Array<{
          delta?: {
            content?: string
            /** DeepSeek/MiniMax native and LiteLLM-normalized thinking stream; OpenRouter uses `reasoning` */
            reasoning_content?: string
            reasoning?: string
            tool_calls?: Array<{
              index: number
              id?: string
              function?: { name?: string; arguments?: string }
            }>
          }
          finish_reason?: string | null
        }>
        error?: { message?: string } | string
      }
    } catch {
      continue
    }
    if (event.error) throw new Error(sseErrorText(event.error, 'Model stream error'))
    const choice = event.choices?.[0]
    if (!choice) continue
    const reasoning = choice.delta?.reasoning_content ?? choice.delta?.reasoning
    if (typeof reasoning === 'string' && reasoning) cb.onReasoningDelta?.(reasoning)
    if (choice.delta?.content) {
      emitted = true
      cb.onDelta(choice.delta.content)
    }
    for (const tc of choice.delta?.tool_calls ?? []) {
      const pending = pendingTools.get(tc.index) ?? {
        id: tc.id ?? crypto.randomUUID(),
        name: '',
        json: '',
      }
      if (tc.id) pending.id = tc.id
      // Spec-compliant servers send the name once, but some local/proxy
      // servers resend the full (or growing) name on every delta; naive
      // concatenation then yields "read_fileread_file" and every call fails
      // as an unknown tool. Treat a delta that extends the accumulated name
      // as a resend, anything else as a fragment to append.
      if (tc.function?.name) {
        pending.name = tc.function.name.startsWith(pending.name)
          ? tc.function.name
          : pending.name + tc.function.name
      }
      if (tc.function?.arguments) pending.json += tc.function.arguments
      pendingTools.set(tc.index, pending)
    }
    if (choice.finish_reason) {
      sawFinish = true
      if (choice.finish_reason === 'length') stopReason = 'max_tokens'
      else if (choice.finish_reason !== 'stop' && choice.finish_reason !== 'tool_calls') {
        abnormalFinish = choice.finish_reason
      }
      flushTools()
    }
  }
  flushTools()
  // e.g. finish_reason=content_filter with no output, or a stream with no
  // message framing at all (gateway soft-failure) — surface both instead of an
  // empty success; a genuine empty turn still carries finish_reason=stop
  if (!emitted && abnormalFinish) {
    throw new Error(`The model returned no content (finish_reason=${abnormalFinish})`)
  }
  if (!emitted && !sawFinish) {
    throw new Error('The model returned no content (empty stream)')
  }
  if (stopReason) cb.onStopReason?.(stopReason)
}

export async function chatOpenAiCompatible(
  wd: StreamWatchdog,
  baseUrl: string,
  config: AiProviderConfig,
  system: string,
  user: string,
  options: OpenAiRequestOptions = {},
): Promise<AiChatResponse> {
  const response = await aiFetch(`${baseUrl.replace(/\/$/, '')}/chat/completions`, {
    method: 'POST',
    signal: wd.signal,
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${config.apiKey}`,
      ...gensparkAttributionHeaders(baseUrl),
    },
    body: JSON.stringify({
      model: config.model,
      messages: [
        { role: 'system', content: system },
        { role: 'user', content: user },
      ],
      ...(options.omitTemperature ? {} : { temperature: 0.3 }),
      ...options.bodyExtras,
    }),
  })
  wd.touch()
  if (!response.ok) {
    return { ok: false, error: `HTTP ${response.status}: ${httpBodyDetail(await response.text())}` }
  }
  const json = (await response.json()) as { choices?: Array<{ message?: { content?: string } }> }
  const content = json.choices?.[0]?.message?.content
  if (!content) return { ok: false, error: 'AI returned an empty response' }
  return { ok: true, content }
}
