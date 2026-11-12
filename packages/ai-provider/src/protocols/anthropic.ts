import type { AgentMessage, AgentToolCall, AgentToolDef } from '@genoffice/agent-core'
import { aiFetch } from '../fetch'
import { httpBodyDetail } from '../http-error'
import {
  GENSPARK_LLM_BASE_URLS,
  gensparkAttributionHeaders,
  opencodeSessionHeaders,
} from '../providers'
import type { AiChatResponse, AiProviderConfig, AiTokenUsage } from '../types'
import { createStreamWatchdog, type StreamWatchdog } from '../watchdog'
import {
  endpointUrl,
  jsonBodyInsteadOfSse,
  parseToolInput,
  readCappedResponseText,
  sseErrorText,
  sseDataEvents,
  throwIfCreditsNotice,
  throwIfToolCountOverBudget,
  throwIfToolJsonOverBudget,
  usageCount,
  type StreamCallbacks,
} from './shared'

export const ANTHROPIC_BASE_URL = 'https://api.anthropic.com'

const EPHEMERAL_CACHE = { type: 'ephemeral' }

/**
 * Prompt caching on the Messages API is opt-in per content block; without the
 * markers every turn of a long tool loop re-bills the whole conversation at
 * the uncached input price (measured on a multi-day editing session: a ~0
 * cache-hit rate). Markers go only to endpoints known to implement them — the
 * official API and the Genspark proxy in front of it; a user-hosted Messages
 * endpoint may reject fields it does not know.
 */
function supportsPromptCache(baseUrl: string): boolean {
  // a stored official URL arrives normalized with a trailing slash
  const base = baseUrl.replace(/\/+$/, '')
  return base === ANTHROPIC_BASE_URL || base === GENSPARK_LLM_BASE_URLS.anthropic
}

/** The provider-wire message shape produced by anthropicMessages */
interface WireMessage {
  role: string
  content: string | Array<Record<string, unknown>>
}

/**
 * Cache breakpoint on the final content block (plain-string user text is
 * upgraded to a text block). cache_control annotations are not part of prefix
 * identity, so the marker moving forward each turn still hits the entry the
 * previous turn wrote.
 */
function markCacheBreakpoint(msg: WireMessage): void {
  if (typeof msg.content === 'string') {
    // an empty text block would be rejected; an empty message cannot anchor a cache entry anyway
    if (!msg.content) return
    msg.content = [{ type: 'text', text: msg.content, cache_control: EPHEMERAL_CACHE }]
    return
  }
  const last = msg.content.at(-1)
  if (last) last['cache_control'] = EPHEMERAL_CACHE
}

function anthropicMessages(messages: AgentMessage[], cache = false): unknown[] {
  const wire = messages.map((m) => {
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
  if (cache) {
    // Last two messages, so the hit-lookup from this turn's marker always has
    // the previous turn's entry within its walk-back window even when a turn
    // appended many blocks (parallel tool results).
    for (const msg of wire.slice(-2)) markCacheBreakpoint(msg as WireMessage)
  }
  return wire
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
    usage?: AnthropicUsage
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
      // A complete JSON body carries the whole turn at once, so the per-turn tool
      // budget of the streamed path has to be applied here as well
      throwIfToolCountOverBudget(toolCalls.length + 1, 'anthropic')
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
  const jsonUsage: AiTokenUsage = {}
  mergeAnthropicUsage(jsonUsage, msg.usage)
  if (jsonUsage.inputTokens !== undefined || jsonUsage.outputTokens !== undefined)
    cb.onUsage?.(jsonUsage)
}

interface AnthropicUsage {
  input_tokens?: unknown
  output_tokens?: unknown
  cache_read_input_tokens?: unknown
  cache_creation_input_tokens?: unknown
}

/** Merge an Anthropic usage object into an accumulator (input arrives on message_start,
 * the running output total on message_delta; stream events and JSON bodies share the shape). */
function mergeAnthropicUsage(target: AiTokenUsage, usage: AnthropicUsage | undefined): void {
  if (!usage) return
  // input_tokens excludes the cached prefix; the billed input is all three together
  const fresh = usageCount(usage.input_tokens)
  const read = usageCount(usage.cache_read_input_tokens)
  const written = usageCount(usage.cache_creation_input_tokens)
  const cached = (read ?? 0) + (written ?? 0)
  const input = fresh === undefined ? (cached > 0 ? cached : undefined) : fresh + cached
  const output = usageCount(usage.output_tokens)
  if (input !== undefined) target.inputTokens = input
  if (output !== undefined) target.outputTokens = output
  if (read !== undefined) target.cacheReadTokens = read
  if (written !== undefined) target.cacheWriteTokens = written
}

/** Per-endpoint request shaping resolved from the provider registry. */
export interface AnthropicRequestOptions {
  /** static headers a user-hosted Messages endpoint needs (merged last, so they win) */
  headers?: Record<string, string> | undefined
  /** stream idle budget override for slow local servers */
  idleTimeoutMs?: number | undefined
  /** adaptive-thinking record for the families that take an effort (4.6+);
   * absent = no thinking block, the vendor's own default. The record carries
   * no level: an `effort` key inside `thinking` is rejected as an invalid
   * request. */
  thinking?: { type: 'adaptive' } | undefined
  /** the level, as the top-level `output_config.effort` that accompanies the
   * adaptive record; absent = the vendor default */
  effort?: string | undefined
}

export async function streamAnthropic(
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
  baseUrl = ANTHROPIC_BASE_URL,
  options: AnthropicRequestOptions = {},
): Promise<void> {
  const wd = createStreamWatchdog(cb.signal, undefined, options.idleTimeoutMs)
  return wd.guard(() =>
    anthropicTurn(config, system, messages, tools, maxTokens, cb, baseUrl, wd, options),
  )
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
  options: AnthropicRequestOptions = {},
): Promise<void> {
  const onBytes = () => {
    wd.touch()
    cb.onActivity?.()
  }
  const cache = supportsPromptCache(baseUrl)
  let response: Response
  try {
    response = await aiFetch(endpointUrl(baseUrl, 'v1/messages'), {
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
        ...opencodeSessionHeaders(baseUrl, cb.sessionId),
        // a user-hosted endpoint may route or authenticate on static headers
        ...options.headers,
      },
      body: JSON.stringify({
        model: config.model,
        max_tokens: maxTokens,
        ...(options.thinking ? { thinking: options.thinking } : {}),
        ...(options.effort ? { output_config: { effort: options.effort } } : {}),
        // the system block's marker also covers the tool definitions ahead of it in the cache prefix
        system:
          cache && system
            ? [{ type: 'text', text: system, cache_control: EPHEMERAL_CACHE }]
            : system,
        messages: anthropicMessages(messages, cache),
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
    throw new Error(
      `Claude HTTP ${response.status}: ${httpBodyDetail(await readCappedResponseText(response, onBytes))}`,
    )
  }
  const jsonBody = await jsonBodyInsteadOfSse(response, onBytes)
  if (jsonBody !== null) {
    throwIfCreditsNotice(jsonBody)
    return emitAnthropicJsonMessage(jsonBody, cb)
  }
  // tool_use inputs stream as partial JSON per content block
  const pendingTools = new Map<number, { id: string; name: string; json: string }>()
  // Some gateways omit the optional block index on delta/stop events; track the
  // last started block so parallel tools don't cross-wire into index 0
  let currentToolIndex = 0
  // emission deferred to stream end: message_delta's stop_reason arrives after all
  // blocks, and a max_tokens stop must mark the last (cut-off) tool call as truncated
  const completedTools: AgentToolCall[] = []
  let stopReason: string | undefined
  let emitted = false
  const turnUsage: AiTokenUsage = {}
  for await (const sse of sseDataEvents(response.body, onBytes)) {
    // A truncated frame or a non-JSON keep-alive from a proxy skips that event
    // rather than killing the turn.
    if (sse.json === undefined) continue
    const event = sse.json as {
      type?: string
      index?: number
      content_block?: { type?: string; id?: string; name?: string }
      delta?: { type?: string; text?: string; partial_json?: string; stop_reason?: string }
      message?: { usage?: AnthropicUsage }
      usage?: AnthropicUsage
      error?: { message?: string } | string
    }
    if (event.type === 'message_start') {
      mergeAnthropicUsage(turnUsage, event.message?.usage)
    } else if (event.type === 'content_block_start' && event.content_block?.type === 'tool_use') {
      const toolIndex = event.index ?? 0
      currentToolIndex = toolIndex
      if (!pendingTools.has(toolIndex)) {
        throwIfToolCountOverBudget(pendingTools.size + completedTools.length + 1, 'anthropic')
      }
      pendingTools.set(toolIndex, {
        id: event.content_block.id ?? crypto.randomUUID(),
        name: event.content_block.name ?? '',
        json: '',
      })
    } else if (event.type === 'content_block_delta') {
      if (event.delta?.type === 'text_delta' && event.delta.text) {
        emitted = true
        cb.onDelta(event.delta.text)
      } else if (event.delta?.type === 'input_json_delta') {
        const pending = pendingTools.get(event.index ?? currentToolIndex)
        if (pending) {
          pending.json += event.delta.partial_json ?? ''
          throwIfToolJsonOverBudget(pending.json.length, 'anthropic')
        }
      }
    } else if (event.type === 'content_block_stop') {
      const stopIndex = event.index ?? currentToolIndex
      const pending = pendingTools.get(stopIndex)
      if (pending) {
        pendingTools.delete(stopIndex)
        // Match the OpenAI route: drop nameless tool calls instead of feeding
        // an empty-name call to the loop (which would always fail as unknown)
        if (pending.name) {
          const { input, error } = parseToolInput(pending.json)
          completedTools.push({ id: pending.id, name: pending.name, input, inputError: error })
        }
      }
    } else if (event.type === 'message_delta') {
      if (event.delta?.stop_reason) stopReason = event.delta.stop_reason
      // message_delta's usage.output_tokens is the running total for the message;
      // it supersedes anything message_start reported for the output side
      mergeAnthropicUsage(turnUsage, event.usage)
    } else if (event.type === 'error' || event.error) {
      // also catches gateway errors delivered in a non-Anthropic shape (no `type` field)
      throw new Error(sseErrorText(event.error, 'Claude stream error'))
    }
  }
  // Buffered tool arguments can take minutes; a gateway dropping the connection
  // meanwhile is a billed in-progress turn, not the replayable empty stream below.
  if (pendingTools.size > 0 && !stopReason) {
    const received = [...pendingTools.values()].reduce((n, p) => n + p.json.length, 0)
    throw new Error(
      `Claude stream closed while sending tool arguments (${received} chars received); the connection was dropped. ` +
        'If this recurs on a large request (e.g. generating a whole document), ask for the output in several smaller parts.',
    )
  }
  // A max_tokens stop can cut a tool_use block before its content_block_stop: emit
  // the partial call as truncated rather than dropping it and answering "done".
  for (const pending of pendingTools.values()) {
    const { input, error } = parseToolInput(pending.json)
    completedTools.push({
      id: pending.id,
      name: pending.name,
      input,
      inputError: error,
      truncated: true,
    })
  }
  pendingTools.clear()
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
  if (!stopReason) {
    throw new Error('Claude stream ended before a stop_reason')
  }
  if (stopReason) cb.onStopReason?.(stopReason)
  if (turnUsage.inputTokens !== undefined || turnUsage.outputTokens !== undefined)
    cb.onUsage?.(turnUsage)
}

export async function chatAnthropic(
  wd: StreamWatchdog,
  config: AiProviderConfig,
  system: string,
  user: string,
  baseUrl = ANTHROPIC_BASE_URL,
  options: AnthropicRequestOptions = {},
): Promise<AiChatResponse> {
  const response = await aiFetch(endpointUrl(baseUrl, 'v1/messages'), {
    method: 'POST',
    signal: wd.signal,
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': config.apiKey,
      'anthropic-version': '2023-06-01',
      // Fetch in the Electron main process goes through Chromium's network stack; this header avoids 403.
      'anthropic-dangerous-direct-browser-access': 'true',
      ...gensparkAttributionHeaders(baseUrl),
      ...opencodeSessionHeaders(baseUrl),
      ...options.headers,
    },
    body: JSON.stringify({
      model: config.model,
      max_tokens: 8192,
      ...(options.thinking ? { thinking: options.thinking } : {}),
      ...(options.effort ? { output_config: { effort: options.effort } } : {}),
      system,
      messages: [{ role: 'user', content: user }],
    }),
  })
  wd.touch()
  if (!response.ok) {
    return {
      ok: false,
      error: `Claude HTTP ${response.status}: ${httpBodyDetail(await readCappedResponseText(response, () => wd.touch()))}`,
    }
  }
  // A 200 with an HTML shell / empty / truncated body (gateway soft-failure)
  // would make response.json() throw; return ok:false instead of leaking a
  // raw SyntaxError to the caller.
  const bodyText = await readCappedResponseText(response, () => wd.touch())
  let json: { content?: Array<{ type: string; text?: string }> }
  try {
    json = JSON.parse(bodyText) as { content?: Array<{ type: string; text?: string }> }
  } catch {
    return {
      ok: false,
      error: `Claude returned a non-JSON response: ${httpBodyDetail(bodyText)}`,
    }
  }
  const content = json.content
    ?.filter((c) => c.type === 'text')
    .map((c) => c.text ?? '')
    .join('')
  if (!content) return { ok: false, error: 'Claude returned an empty response' }
  return { ok: true, content }
}
