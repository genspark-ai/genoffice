import type { AgentMessage, AgentToolCall, AgentToolDef } from '@genoffice/agent-core'
import { aiFetch } from '../fetch'
import { httpBodyDetail } from '../http-error'
import { gensparkAttributionHeaders } from '../providers'
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

export const ANTHROPIC_BASE_URL = 'https://api.anthropic.com'

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
  baseUrl = ANTHROPIC_BASE_URL,
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

export async function chatAnthropic(
  wd: StreamWatchdog,
  config: AiProviderConfig,
  system: string,
  user: string,
  baseUrl = ANTHROPIC_BASE_URL,
): Promise<AiChatResponse> {
  const response = await aiFetch(`${baseUrl.replace(/\/$/, '')}/v1/messages`, {
    method: 'POST',
    signal: wd.signal,
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': config.apiKey,
      'anthropic-version': '2023-06-01',
      // Fetch in the Electron main process goes through Chromium's network stack; this header avoids 403.
      'anthropic-dangerous-direct-browser-access': 'true',
      ...gensparkAttributionHeaders(baseUrl),
    },
    body: JSON.stringify({
      model: config.model,
      max_tokens: 8192,
      system,
      messages: [{ role: 'user', content: user }],
    }),
  })
  wd.touch()
  if (!response.ok) {
    return {
      ok: false,
      error: `Claude HTTP ${response.status}: ${httpBodyDetail(await response.text())}`,
    }
  }
  const json = (await response.json()) as { content?: Array<{ type: string; text?: string }> }
  const content = json.content
    ?.filter((c) => c.type === 'text')
    .map((c) => c.text ?? '')
    .join('')
  if (!content) return { ok: false, error: 'Claude returned an empty response' }
  return { ok: true, content }
}
