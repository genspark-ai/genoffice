import type { AgentMessage, AgentToolCall, AgentToolDef } from '@genoffice/agent-core'
import { aiFetch } from './fetch'
import { httpBodyDetail } from './http-error'
import { GENSPARK_LLM_BASE_URLS, gensparkAttributionHeaders } from './providers'
import type { AiProviderConfig, AiProviderId } from './types'
import { createStreamWatchdog, type StreamWatchdog } from './watchdog'

// ---- streaming (SSE line splitting shared by all providers) ----

export async function* sseLines(
  body: NodeJS.ReadableStream | ReadableStream<Uint8Array>,
  onBytes?: () => void,
): AsyncGenerator<string> {
  const decoder = new TextDecoder()
  let buffer = ''
  const stream = body as ReadableStream<Uint8Array>
  const reader = stream.getReader()
  while (true) {
    const { done, value } = await reader.read()
    if (done) break
    onBytes?.()
    buffer += decoder.decode(value, { stream: true })
    const lines = buffer.split('\n')
    buffer = lines.pop() ?? ''
    for (const line of lines) yield line
  }
  if (buffer) yield buffer
}

export interface StreamCallbacks {
  onDelta: (text: string) => void
  onToolCall: (call: AgentToolCall) => void
  /** normalized stop reason ('max_tokens' when the output was cut off by the token limit) */
  onStopReason?: (reason: string) => void
  /** bytes arrived on the wire (fires per network chunk, including SSE pings; used for keepalive) */
  onActivity?: () => void
  signal: AbortSignal
}

/**
 * Models occasionally emit unescaped " inside string values (e.g. English quotes in Chinese copy).
 * Single-pass scan: a " inside a string whose next non-whitespace char is not structural gets escaped.
 */
function repairUnescapedQuotes(json: string): string {
  let out = ''
  let inStr = false
  for (let i = 0; i < json.length; i++) {
    const c = json[i]!
    if (!inStr) {
      if (c === '"') inStr = true
      out += c
      continue
    }
    if (c === '\\') {
      out += c + (json[++i] ?? '')
      continue
    }
    if (c === '"') {
      let j = i + 1
      while (j < json.length && ' \n\r\t'.includes(json[j]!)) j++
      const next = json[j]
      if (next === undefined || ',}]:'.includes(next)) {
        inStr = false
        out += c
      } else {
        out += '\\"'
      }
      continue
    }
    out += c
  }
  return out
}

/**
 * Gateways can report failures (quota exhausted, moderation, upstream errors) inside a
 * 200 SSE stream, in shapes that don't match the provider protocol (e.g. an OpenAI-style
 * `{"error": ...}` event on the Anthropic route). Extract a readable message so these
 * surface as real errors instead of dissolving into an empty "successful" turn.
 */
function sseErrorText(error: unknown, fallback: string): string {
  if (typeof error === 'string' && error) return error
  if (error && typeof error === 'object') {
    const message = (error as { message?: unknown }).message
    if (typeof message === 'string' && message) return message
    try {
      return JSON.stringify(error)
    } catch {
      /* circular or otherwise unserializable — use the fallback */
    }
  }
  return fallback
}

/**
 * Gateways can answer a `stream: true` request with a complete non-SSE JSON body —
 * observed on the Genspark Anthropic route when credits are exhausted (HTTP 200,
 * Content-Type: application/json, the notice text inside a regular message). The SSE
 * parser would find no `data:` lines in such a body and dissolve it into an empty
 * "successful" turn. Returns the body text when that happens, else null.
 */
async function jsonBodyInsteadOfSse(response: Response): Promise<string | null> {
  const contentType = response.headers.get('content-type') ?? ''
  return contentType.includes('application/json') ? await response.text() : null
}

/**
 * A non-SSE JSON reply whose text is the gateway's credits-exhausted notice
 * (Genspark: "Your Genspark credits have been exhausted…") surfaces as a typed
 * error so the apps show a localized "top up" message (errorCode 'credits')
 * instead of the English notice as a normal assistant reply.
 */
export class AiCreditsError extends Error {
  constructor(notice: string) {
    super(notice)
    this.name = 'AiCreditsError'
  }
}

function creditsNoticeText(value: unknown): string | null {
  if (typeof value === 'string') {
    const t = value.toLowerCase()
    const credits =
      t.includes('genspark.ai/pricing') ||
      (t.includes('credit') && (t.includes('exhausted') || t.includes('insufficient')))
    return credits ? value : null
  }
  if (Array.isArray(value) || (value && typeof value === 'object')) {
    for (const v of Object.values(value)) {
      const hit = creditsNoticeText(v)
      if (hit) return hit
    }
  }
  return null
}

function throwIfCreditsNotice(bodyText: string): void {
  let parsed: unknown
  try {
    parsed = JSON.parse(bodyText)
  } catch {
    return // unparseable bodies are the emit helpers' problem
  }
  const notice = creditsNoticeText(parsed)
  if (notice) throw new AiCreditsError(notice)
}

/** Don't throw on parse failure (it would kill the whole stream); return error so the loop feeds it back for retry */
function parseToolInput(json: string): { input: Record<string, unknown>; error?: string } {
  if (!json.trim()) return { input: {} }
  try {
    return { input: JSON.parse(json) as Record<string, unknown> }
  } catch (e) {
    try {
      return { input: JSON.parse(repairUnescapedQuotes(json)) as Record<string, unknown> }
    } catch {
      const msg = e instanceof Error ? e.message : String(e)
      return { input: {}, error: `${msg}; raw: ${json.slice(0, 500)}` }
    }
  }
}

// ---- Anthropic ----

function anthropicMessages(messages: AgentMessage[]): unknown[] {
  return messages.map((m) => {
    if (m.role === 'user') {
      // Keep plain-text content as a string; only upgrade to a content block array when images are present
      if (!m.images?.length) return { role: 'user', content: m.text }
      return {
        role: 'user',
        content: [
          ...(m.text ? [{ type: 'text', text: m.text }] : []),
          ...m.images.map((img) => ({
            type: 'image',
            source: { type: 'base64', media_type: img.mime, data: img.base64 },
          })),
        ],
      }
    }
    if (m.role === 'assistant') {
      const content: unknown[] = []
      if (m.text) content.push({ type: 'text', text: m.text })
      for (const call of m.toolCalls ?? []) {
        content.push({ type: 'tool_use', id: call.id, name: call.name, input: call.input })
      }
      // Anthropic rejects empty content arrays; a prior empty terminal turn
      // (tool work with no prose) would otherwise poison every follow-up.
      if (content.length === 0) content.push({ type: 'text', text: '(no content)' })
      return { role: 'assistant', content }
    }
    // tool results travel back as a user message of tool_result blocks
    return {
      role: 'user',
      content: m.results.map((r) => ({
        type: 'tool_result',
        tool_use_id: r.id,
        content: r.output,
        ...(r.isError ? { is_error: true } : {}),
      })),
    }
  })
}

/** Emits a complete (non-streamed) Anthropic message delivered as a plain JSON body. */
function emitAnthropicJsonMessage(bodyText: string, cb: StreamCallbacks): void {
  let msg: {
    content?: Array<{
      type?: string
      text?: string
      id?: string
      name?: string
      input?: Record<string, unknown>
    }>
    stop_reason?: string
    error?: { message?: string } | string
  }
  try {
    msg = JSON.parse(bodyText) as typeof msg
  } catch {
    throw new Error(`Claude returned an unparseable JSON body: ${httpBodyDetail(bodyText)}`)
  }
  if (msg.error) throw new Error(sseErrorText(msg.error, 'Claude error'))
  let emitted = false
  const toolCalls: AgentToolCall[] = []
  for (const block of msg.content ?? []) {
    if (block.type === 'text' && block.text) {
      emitted = true
      cb.onDelta(block.text)
    } else if (block.type === 'tool_use' && block.name) {
      emitted = true
      toolCalls.push({
        id: block.id ?? crypto.randomUUID(),
        name: block.name,
        input: block.input ?? {},
      })
    }
  }
  // a max_tokens stop may have cut off the last tool call's arguments
  const lastTool = toolCalls.at(-1)
  if (msg.stop_reason === 'max_tokens' && lastTool) lastTool.truncated = true
  for (const call of toolCalls) cb.onToolCall(call)
  if (!emitted) throw new Error(`Claude returned no content: ${httpBodyDetail(bodyText)}`)
  if (msg.stop_reason) cb.onStopReason?.(msg.stop_reason)
}

export async function streamAnthropic(
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
  baseUrl = 'https://api.anthropic.com',
): Promise<void> {
  const wd = createStreamWatchdog(cb.signal)
  return wd.guard(() => anthropicTurn(config, system, messages, tools, maxTokens, cb, baseUrl, wd))
}

async function anthropicTurn(
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
  baseUrl: string,
  wd: StreamWatchdog,
): Promise<void> {
  const onBytes = () => {
    wd.touch()
    cb.onActivity?.()
  }
  let response: Response
  try {
    response = await aiFetch(`${baseUrl.replace(/\/$/, '')}/v1/messages`, {
      method: 'POST',
      signal: wd.signal,
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': config.apiKey,
        'anthropic-version': '2023-06-01',
        // Renderer fetches and the net.fetch rescue path go through Chromium's network stack,
        // which adds browser-semantics headers; Anthropic rejects those with 403 "Request not
        // allowed". This header is the official opt-in for browser/Electron environments.
        'anthropic-dangerous-direct-browser-access': 'true',
        ...gensparkAttributionHeaders(baseUrl),
      },
      body: JSON.stringify({
        model: config.model,
        max_tokens: maxTokens,
        system,
        messages: anthropicMessages(messages),
        ...(tools.length > 0
          ? {
              tools: tools.map((t) => ({
                name: t.name,
                description: t.description,
                input_schema: t.inputSchema,
              })),
            }
          : {}),
        stream: true,
      }),
    })
  } catch (e) {
    // When fetch fails in the Electron main process, the real reason lives in `cause`
    const err = e as { message?: unknown; cause?: { code?: unknown; message?: unknown } } | null
    const causeText = err?.cause
      ? ` cause=${String(err.cause.code || err.cause.message || err.cause)}`
      : ''
    throw new Error(`Claude fetch failed: ${err?.message || String(e)}${causeText}`, { cause: e })
  }
  // headers arrived: ping the renderer watchdog too, or a slow first chunk could trip it
  onBytes()
  if (!response.ok || !response.body) {
    throw new Error(`Claude HTTP ${response.status}: ${httpBodyDetail(await response.text())}`)
  }
  const jsonBody = await jsonBodyInsteadOfSse(response)
  if (jsonBody !== null) {
    throwIfCreditsNotice(jsonBody)
    return emitAnthropicJsonMessage(jsonBody, cb)
  }
  // tool_use inputs stream as partial JSON per content block
  const pendingTools = new Map<number, { id: string; name: string; json: string }>()
  // emission deferred to stream end: message_delta's stop_reason arrives after all
  // blocks, and a max_tokens stop must mark the last (cut-off) tool call as truncated
  const completedTools: AgentToolCall[] = []
  let stopReason: string | undefined
  let emitted = false
  for await (const line of sseLines(response.body, onBytes)) {
    if (!line.startsWith('data:')) continue
    const payload = line.slice(5).trim()
    if (!payload) continue
    const event = JSON.parse(payload) as {
      type?: string
      index?: number
      content_block?: { type?: string; id?: string; name?: string }
      delta?: { type?: string; text?: string; partial_json?: string; stop_reason?: string }
      error?: { message?: string } | string
    }
    if (event.type === 'content_block_start' && event.content_block?.type === 'tool_use') {
      pendingTools.set(event.index ?? 0, {
        id: event.content_block.id ?? crypto.randomUUID(),
        name: event.content_block.name ?? '',
        json: '',
      })
    } else if (event.type === 'content_block_delta') {
      if (event.delta?.type === 'text_delta' && event.delta.text) {
        emitted = true
        cb.onDelta(event.delta.text)
      } else if (event.delta?.type === 'input_json_delta') {
        const pending = pendingTools.get(event.index ?? 0)
        if (pending) pending.json += event.delta.partial_json ?? ''
      }
    } else if (event.type === 'content_block_stop') {
      const pending = pendingTools.get(event.index ?? 0)
      if (pending) {
        pendingTools.delete(event.index ?? 0)
        const { input, error } = parseToolInput(pending.json)
        completedTools.push({ id: pending.id, name: pending.name, input, inputError: error })
      }
    } else if (event.type === 'message_delta') {
      if (event.delta?.stop_reason) stopReason = event.delta.stop_reason
    } else if (event.type === 'error' || event.error) {
      // also catches gateway errors delivered in a non-Anthropic shape (no `type` field)
      throw new Error(sseErrorText(event.error, 'Claude stream error'))
    }
  }
  const lastTool = completedTools.at(-1)
  if (stopReason === 'max_tokens' && lastTool) lastTool.truncated = true
  for (const call of completedTools) cb.onToolCall(call)
  // A stream with no content AND no message framing (no stop_reason ever seen)
  // is a gateway soft-failure, not a model turn — surface it instead of letting
  // it dissolve into an empty "successful" turn with no diagnostics. A genuine
  // empty closing turn (common after tool-heavy runs) still carries end_turn.
  // The "(empty stream)" suffix is a contract: app renderers match it to
  // classify the failure as empty output (fail fast, no billed retries).
  if (!emitted && completedTools.length === 0 && !stopReason) {
    throw new Error('Claude returned no content (empty stream)')
  }
  if (stopReason) cb.onStopReason?.(stopReason)
}

// ---- Gemini ----

function geminiContents(messages: AgentMessage[]): unknown[] {
  return messages.map((m) => {
    if (m.role === 'user') {
      if (!m.images?.length) return { role: 'user', parts: [{ text: m.text }] }
      return {
        role: 'user',
        parts: [
          ...(m.text ? [{ text: m.text }] : []),
          ...m.images.map((img) => ({ inline_data: { mime_type: img.mime, data: img.base64 } })),
        ],
      }
    }
    if (m.role === 'assistant') {
      const parts: unknown[] = []
      if (m.text) parts.push({ text: m.text })
      for (const call of m.toolCalls ?? []) {
        parts.push({ functionCall: { name: call.name, args: call.input } })
      }
      // Gemini rejects model turns with empty parts lists.
      if (parts.length === 0) parts.push({ text: '(no content)' })
      return { role: 'model', parts }
    }
    return {
      role: 'user',
      parts: m.results.map((r) => ({
        functionResponse: {
          name: r.name,
          response: r.isError ? { error: r.output } : { result: r.output },
        },
      })),
    }
  })
}

/**
 * Emits a complete (non-streamed) Gemini response delivered as a plain JSON body.
 * `streamGenerateContent` without SSE framing yields an array of chunks; a gateway
 * may also send a single `generateContent`-shaped object — handle both.
 */
function emitGeminiJsonMessage(bodyText: string, cb: StreamCallbacks): void {
  let parsed: unknown
  try {
    parsed = JSON.parse(bodyText)
  } catch {
    throw new Error(`Gemini returned an unparseable JSON body: ${httpBodyDetail(bodyText)}`)
  }
  const events = (Array.isArray(parsed) ? parsed : [parsed]) as Array<{
    candidates?: Array<{
      content?: {
        parts?: Array<{
          text?: string
          functionCall?: { name?: string; args?: Record<string, unknown> }
        }>
      }
      finishReason?: string
    }>
    promptFeedback?: { blockReason?: string }
    error?: { message?: string } | string
  }>
  let emitted = false
  let stopReason: string | undefined
  let abnormalFinish: string | undefined
  for (const event of events) {
    if (event.error) throw new Error(sseErrorText(event.error, 'Gemini error'))
    if (event.promptFeedback?.blockReason) {
      throw new Error(`Gemini blocked the prompt (${event.promptFeedback.blockReason})`)
    }
    const finishReason = event.candidates?.[0]?.finishReason
    if (finishReason === 'MAX_TOKENS') stopReason = 'max_tokens'
    else if (finishReason && finishReason !== 'STOP') abnormalFinish = finishReason
    for (const part of event.candidates?.[0]?.content?.parts ?? []) {
      if (part.text) {
        emitted = true
        cb.onDelta(part.text)
      }
      if (part.functionCall?.name) {
        emitted = true
        cb.onToolCall({
          id: crypto.randomUUID(),
          name: part.functionCall.name,
          input: part.functionCall.args ?? {},
        })
      }
    }
  }
  if (!emitted) {
    throw new Error(
      abnormalFinish
        ? `Gemini returned no content (finishReason=${abnormalFinish})`
        : `Gemini returned no content: ${httpBodyDetail(bodyText)}`,
    )
  }
  if (stopReason) cb.onStopReason?.(stopReason)
}

export async function streamGemini(
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
  baseUrl = 'https://generativelanguage.googleapis.com/v1beta',
): Promise<void> {
  const wd = createStreamWatchdog(cb.signal)
  return wd.guard(() => geminiTurn(config, system, messages, tools, maxTokens, cb, baseUrl, wd))
}

async function geminiTurn(
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
  baseUrl: string,
  wd: StreamWatchdog,
): Promise<void> {
  const onBytes = () => {
    wd.touch()
    cb.onActivity?.()
  }
  const url = `${baseUrl.replace(/\/$/, '')}/models/${config.model}:streamGenerateContent?alt=sse`
  const response = await aiFetch(url, {
    method: 'POST',
    signal: wd.signal,
    headers: {
      'Content-Type': 'application/json',
      'x-goog-api-key': config.apiKey,
      ...gensparkAttributionHeaders(baseUrl),
    },
    body: JSON.stringify({
      systemInstruction: { parts: [{ text: system }] },
      contents: geminiContents(messages),
      ...(tools.length > 0
        ? {
            tools: [
              {
                functionDeclarations: tools.map((t) => ({
                  name: t.name,
                  description: t.description,
                  parameters: t.inputSchema,
                })),
              },
            ],
          }
        : {}),
      generationConfig: { temperature: 0.3, maxOutputTokens: maxTokens },
    }),
  })
  // headers arrived: ping the renderer watchdog too, or a slow first chunk could trip it
  onBytes()
  if (!response.ok || !response.body) {
    throw new Error(`Gemini HTTP ${response.status}: ${httpBodyDetail(await response.text())}`)
  }
  const jsonBody = await jsonBodyInsteadOfSse(response)
  if (jsonBody !== null) {
    throwIfCreditsNotice(jsonBody)
    return emitGeminiJsonMessage(jsonBody, cb)
  }
  let stopReason: string | undefined
  let abnormalFinish: string | undefined
  let sawFinish = false
  let emitted = false
  for await (const line of sseLines(response.body, onBytes)) {
    if (!line.startsWith('data:')) continue
    const payload = line.slice(5).trim()
    if (!payload) continue
    const event = JSON.parse(payload) as {
      candidates?: Array<{
        content?: {
          parts?: Array<{
            text?: string
            functionCall?: { name?: string; args?: Record<string, unknown> }
          }>
        }
        finishReason?: string
      }>
      promptFeedback?: { blockReason?: string }
      error?: { message?: string } | string
    }
    if (event.error) throw new Error(sseErrorText(event.error, 'Gemini stream error'))
    if (event.promptFeedback?.blockReason) {
      throw new Error(`Gemini blocked the prompt (${event.promptFeedback.blockReason})`)
    }
    const finishReason = event.candidates?.[0]?.finishReason
    if (finishReason) sawFinish = true
    if (finishReason === 'MAX_TOKENS') stopReason = 'max_tokens'
    else if (finishReason && finishReason !== 'STOP') abnormalFinish = finishReason
    for (const part of event.candidates?.[0]?.content?.parts ?? []) {
      if (part.text) {
        emitted = true
        cb.onDelta(part.text)
      }
      // Gemini emits function calls whole, never as partial JSON
      if (part.functionCall?.name) {
        emitted = true
        cb.onToolCall({
          id: crypto.randomUUID(),
          name: part.functionCall.name,
          input: part.functionCall.args ?? {},
        })
      }
    }
  }
  // A safety/recitation stop that produced nothing, or a stream with no message
  // framing at all (gateway soft-failure), would otherwise look like an empty
  // success; a genuine empty turn still carries finishReason=STOP and passes
  if (!emitted && abnormalFinish) {
    throw new Error(`Gemini returned no content (finishReason=${abnormalFinish})`)
  }
  if (!emitted && !sawFinish) {
    throw new Error('Gemini returned no content (empty stream)')
  }
  if (stopReason) cb.onStopReason?.(stopReason)
}

// ---- OpenAI-compatible (openai / deepseek / custom) ----

function openAiMessages(system: string, messages: AgentMessage[]): unknown[] {
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

export async function streamOpenAiCompatible(
  baseUrl: string,
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
): Promise<void> {
  const wd = createStreamWatchdog(cb.signal)
  return wd.guard(() =>
    openAiCompatibleTurn(baseUrl, config, system, messages, tools, maxTokens, cb, wd),
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
      max_tokens: maxTokens,
      messages: openAiMessages(system, messages),
      ...(tools.length > 0
        ? {
            tools: tools.map((t) => ({
              type: 'function',
              function: { name: t.name, description: t.description, parameters: t.inputSchema },
            })),
          }
        : {}),
      temperature: 0.3,
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
    const event = JSON.parse(payload) as {
      choices?: Array<{
        delta?: {
          content?: string
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
    if (event.error) throw new Error(sseErrorText(event.error, 'Model stream error'))
    const choice = event.choices?.[0]
    if (!choice) continue
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
      if (tc.function?.name) pending.name += tc.function.name
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

const OPENAI_COMPATIBLE_BASE_URLS: Partial<Record<AiProviderId, string>> = {
  deepseek: 'https://api.deepseek.com/v1',
  openai: 'https://api.openai.com/v1',
}

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
  switch (provider) {
    case 'genspark':
      // The proxy exposes three protocol-specific endpoints; route by model id prefix: claude uses
      // the Anthropic protocol (preserves image input fidelity), gemini uses Gemini, rest OpenAI-compatible
      if (config.model.startsWith('claude')) {
        return streamAnthropic(
          config,
          system,
          messages,
          tools,
          maxTokens,
          cb,
          GENSPARK_LLM_BASE_URLS.anthropic,
        )
      }
      if (config.model.startsWith('gemini')) {
        return streamGemini(
          config,
          system,
          messages,
          tools,
          maxTokens,
          cb,
          GENSPARK_LLM_BASE_URLS.gemini,
        )
      }
      return streamOpenAiCompatible(
        GENSPARK_LLM_BASE_URLS.openai,
        config,
        system,
        messages,
        tools,
        maxTokens,
        cb,
      )
    case 'anthropic':
      return streamAnthropic(config, system, messages, tools, maxTokens, cb)
    case 'gemini':
      return streamGemini(config, system, messages, tools, maxTokens, cb)
    case 'deepseek':
    case 'openai':
      return streamOpenAiCompatible(
        OPENAI_COMPATIBLE_BASE_URLS[provider]!,
        config,
        system,
        messages,
        tools,
        maxTokens,
        cb,
      )
    case 'custom':
      if (!config.baseUrl) throw new Error('A custom provider requires a Base URL')
      return streamOpenAiCompatible(config.baseUrl, config, system, messages, tools, maxTokens, cb)
    default:
      throw new Error(`Unknown provider: ${provider}`)
  }
}
