import type { AgentMessage, AgentToolCall, AgentToolDef } from '@genoffice/agent-core'
import { aiFetch } from '../fetch'
import { httpBodyDetail } from '../http-error'
import type { AiChatResponse, AiProviderConfig, AiTokenUsage } from '../types'
import { createStreamWatchdog, type StreamWatchdog } from '../watchdog'
import {
  endpointUrl,
  jsonBodyInsteadOfSse,
  parseToolInput,
  readCappedResponseText,
  sseDataEvents,
  sseErrorText,
  throwIfCreditsNotice,
  throwIfToolCountOverBudget,
  throwIfToolJsonOverBudget,
  usageCount,
  type StreamCallbacks,
} from './shared'

/**
 * OpenAI Responses API (`POST /v1/responses`). GPT-5.4+ and GPT-6 call tools
 * only here (Chat Completions 400s unless reasoning is off), so this is the
 * route that keeps their reasoning on in the agent loop.
 *
 * Nothing is stored server-side (`store: false`); the reasoning the model did
 * before a tool call comes back as an encrypted item and rides on the turn's
 * first tool call `signature`, the same opaque slot Gemini uses for its thought
 * signature, so the next request can replay it ahead of the function calls.
 */

const REASONING_SIGNATURE_PREFIX = 'oai-rs:'

function reasoningSignature(id: string, encrypted: string): string {
  return REASONING_SIGNATURE_PREFIX + JSON.stringify({ id, encrypted })
}

function parseReasoningSignature(
  signature: string | undefined,
): { id: string; encrypted: string } | null {
  if (!signature?.startsWith(REASONING_SIGNATURE_PREFIX)) return null
  try {
    const parsed = JSON.parse(signature.slice(REASONING_SIGNATURE_PREFIX.length)) as {
      id?: unknown
      encrypted?: unknown
    }
    return typeof parsed.id === 'string' && typeof parsed.encrypted === 'string'
      ? { id: parsed.id, encrypted: parsed.encrypted }
      : null
  } catch {
    return null
  }
}

function responsesInput(messages: AgentMessage[]): unknown[] {
  const out: unknown[] = []
  for (const m of messages) {
    if (m.role === 'user') {
      out.push({
        role: 'user',
        content: [
          ...(m.text || !m.images?.length ? [{ type: 'input_text', text: m.text }] : []),
          ...(m.images ?? []).map((img) => ({
            type: 'input_image',
            image_url: `data:${img.mime};base64,${img.base64}`,
            detail: 'auto',
          })),
        ],
      })
    } else if (m.role === 'assistant') {
      const calls = m.toolCalls ?? []
      const reasoning = parseReasoningSignature(calls[0]?.signature)
      if (reasoning) {
        out.push({
          type: 'reasoning',
          id: reasoning.id,
          encrypted_content: reasoning.encrypted,
          summary: [],
        })
      }
      // an assistant item with no text and no calls is an empty turn the API
      // rejects, so keep the placeholder the Chat Completions path uses
      if (m.text || calls.length === 0) {
        out.push({
          role: 'assistant',
          content: [{ type: 'output_text', text: m.text || '(no content)' }],
        })
      }
      for (const call of calls) {
        out.push({
          type: 'function_call',
          call_id: call.id,
          name: call.name,
          arguments: JSON.stringify(call.input),
        })
      }
    } else {
      for (const r of m.results) {
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

/** Per-endpoint request shaping resolved from the provider registry. */
export interface OpenAiResponsesOptions {
  /** `reasoning.effort` to request; absent = the model's default */
  reasoningEffort?: string | undefined
  /** vendor-specific fields merged into the request body */
  bodyExtras?: Record<string, unknown> | undefined
  headers?: Record<string, string> | undefined
  idleTimeoutMs?: number | undefined
}

function requestBody(
  config: AiProviderConfig,
  system: string,
  input: unknown[],
  tools: AgentToolDef[],
  maxTokens: number,
  options: OpenAiResponsesOptions,
  stream: boolean,
): string {
  return JSON.stringify({
    model: config.model,
    instructions: system,
    input,
    ...(tools.length > 0
      ? {
          tools: tools.map((t) => ({
            type: 'function',
            name: t.name,
            description: t.description,
            parameters: t.inputSchema,
            strict: false,
          })),
        }
      : {}),
    max_output_tokens: maxTokens,
    store: false,
    include: ['reasoning.encrypted_content'],
    reasoning: {
      summary: 'auto',
      ...(options.reasoningEffort ? { effort: options.reasoningEffort } : {}),
    },
    ...options.bodyExtras,
    stream,
  })
}

interface ResponsesItem {
  id?: string
  type?: string
  call_id?: string
  name?: string
  arguments?: string
  encrypted_content?: string
  content?: Array<{ type?: string; text?: string }>
  summary?: Array<{ type?: string; text?: string }>
}

interface ResponsesObject {
  status?: string
  output?: ResponsesItem[]
  incomplete_details?: { reason?: string }
  usage?: { input_tokens?: unknown; output_tokens?: unknown }
  error?: { message?: string } | string | null
}

function usageOf(raw: ResponsesObject['usage']): AiTokenUsage | undefined {
  if (!raw) return undefined
  const usage: AiTokenUsage = {}
  const input = usageCount(raw.input_tokens)
  const output = usageCount(raw.output_tokens)
  if (input !== undefined) usage.inputTokens = input
  if (output !== undefined) usage.outputTokens = output
  return usage.inputTokens !== undefined || usage.outputTokens !== undefined ? usage : undefined
}

/** Turn the completed items of one response into callbacks; shared by the SSE and JSON paths. */
function emitFinishedItems(
  items: ResponsesItem[],
  response: ResponsesObject | undefined,
  cb: StreamCallbacks,
  alreadyStreamed: boolean,
): boolean {
  let emitted = alreadyStreamed
  let reasoning: { id: string; encrypted: string } | undefined
  const toolCalls: AgentToolCall[] = []
  for (const item of items) {
    if (item.type === 'reasoning') {
      if (item.id && item.encrypted_content) {
        reasoning = { id: item.id, encrypted: item.encrypted_content }
      }
      if (!alreadyStreamed) {
        for (const part of item.summary ?? []) {
          if (part.text) {
            emitted = true
            cb.onReasoningDelta?.(part.text)
          }
        }
      }
    } else if (item.type === 'message') {
      if (!alreadyStreamed) {
        for (const part of item.content ?? []) {
          if (part.type === 'output_text' && part.text) {
            emitted = true
            cb.onDelta(part.text)
          }
        }
      }
    } else if (item.type === 'function_call' && item.name) {
      throwIfToolCountOverBudget(toolCalls.length + 1, 'openai-responses')
      const { input, error } = parseToolInput(item.arguments ?? '')
      toolCalls.push({
        id: item.call_id ?? item.id ?? crypto.randomUUID(),
        name: item.name,
        input,
        inputError: error,
      })
    }
  }
  const cutOff = response?.incomplete_details?.reason === 'max_output_tokens'
  const lastTool = toolCalls.at(-1)
  if (cutOff && lastTool) lastTool.truncated = true
  if (reasoning && toolCalls[0]) {
    toolCalls[0].signature = reasoningSignature(reasoning.id, reasoning.encrypted)
  }
  for (const call of toolCalls) {
    emitted = true
    cb.onToolCall(call)
  }
  if (cutOff) cb.onStopReason?.('max_tokens')
  const usage = usageOf(response?.usage)
  if (usage) cb.onUsage?.(usage)
  return emitted
}

function emitJsonResponse(bodyText: string, cb: StreamCallbacks): void {
  let response: ResponsesObject
  try {
    response = JSON.parse(bodyText) as ResponsesObject
  } catch {
    throw new Error(`The model returned an unparseable JSON body: ${httpBodyDetail(bodyText)}`)
  }
  if (response.error) throw new Error(sseErrorText(response.error, 'Model error'))
  const emitted = emitFinishedItems(response.output ?? [], response, cb, false)
  if (!emitted) throw new Error(`The model returned no content: ${httpBodyDetail(bodyText)}`)
}

export async function streamOpenAiResponses(
  baseUrl: string,
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
  options: OpenAiResponsesOptions = {},
): Promise<void> {
  const wd = createStreamWatchdog(cb.signal, undefined, options.idleTimeoutMs)
  return wd.guard(() =>
    responsesTurn(baseUrl, config, system, messages, tools, maxTokens, cb, wd, options),
  )
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
  options: OpenAiResponsesOptions,
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
      ...options.headers,
    },
    body: requestBody(config, system, responsesInput(messages), tools, maxTokens, options, true),
  })
  onBytes()
  if (!response.ok || !response.body) {
    throw new Error(
      `HTTP ${response.status}: ${httpBodyDetail(await readCappedResponseText(response, onBytes))}`,
    )
  }
  const jsonBody = await jsonBodyInsteadOfSse(response, onBytes)
  if (jsonBody !== null) {
    throwIfCreditsNotice(jsonBody)
    return emitJsonResponse(jsonBody, cb)
  }
  // function_call arguments stream per item; the item itself arrives complete
  // in output_item.done, which is what the turn is assembled from
  const argumentChars = new Map<string, number>()
  const finished: ResponsesItem[] = []
  let streamedAny = false
  let completed: ResponsesObject | undefined
  let terminal = false
  for await (const sse of sseDataEvents(response.body, onBytes)) {
    if (sse.raw === '[DONE]') break
    if (sse.json === undefined) continue
    const event = sse.json as {
      type?: string
      delta?: string
      item_id?: string
      item?: ResponsesItem
      response?: ResponsesObject
      error?: { message?: string } | string
      message?: string
    }
    switch (event.type) {
      case 'response.output_text.delta':
        if (event.delta) {
          streamedAny = true
          cb.onDelta(event.delta)
        }
        break
      case 'response.reasoning_summary_text.delta':
        if (event.delta) {
          streamedAny = true
          cb.onReasoningDelta?.(event.delta)
        }
        break
      case 'response.output_item.added':
        if (event.item?.type === 'function_call') {
          throwIfToolCountOverBudget(argumentChars.size + 1, 'openai-responses')
          argumentChars.set(event.item.id ?? String(argumentChars.size), 0)
        }
        break
      case 'response.function_call_arguments.delta': {
        const key = event.item_id ?? ''
        const total = (argumentChars.get(key) ?? 0) + (event.delta?.length ?? 0)
        throwIfToolJsonOverBudget(total, 'openai-responses')
        argumentChars.set(key, total)
        break
      }
      case 'response.output_item.done':
        if (event.item) finished.push(event.item)
        break
      case 'response.completed':
      case 'response.incomplete':
        completed = event.response
        terminal = true
        break
      case 'response.failed':
        throw new Error(
          sseErrorText(event.response?.error ?? 'response.failed', 'Model stream error'),
        )
      case 'error':
        throw new Error(sseErrorText(event.error ?? event.message ?? event, 'Model stream error'))
    }
    if (terminal) break
  }
  if (!terminal) {
    if (finished.some((item) => item.type === 'function_call') || argumentChars.size > 0) {
      throw new Error(
        'The model stream closed while sending tool arguments; the connection was dropped. ' +
          'If this recurs on a large request (e.g. generating a whole document), ask for the output in several smaller parts.',
      )
    }
    if (streamedAny) throw new Error('The model stream ended before response.completed')
  }
  const emitted = emitFinishedItems(finished, completed, cb, streamedAny)
  if (!emitted) {
    throw new Error(
      terminal
        ? `The model returned no content (status=${completed?.status ?? 'unknown'})`
        : 'The model returned no content (empty stream)',
    )
  }
}

export async function chatOpenAiResponses(
  wd: StreamWatchdog,
  baseUrl: string,
  config: AiProviderConfig,
  system: string,
  user: string,
  options: OpenAiResponsesOptions = {},
): Promise<AiChatResponse> {
  const response = await aiFetch(endpointUrl(baseUrl, 'responses'), {
    method: 'POST',
    signal: wd.signal,
    headers: {
      'Content-Type': 'application/json',
      ...(config.apiKey ? { Authorization: `Bearer ${config.apiKey}` } : {}),
      ...options.headers,
    },
    body: requestBody(
      config,
      system,
      responsesInput([{ role: 'user', text: user }]),
      [],
      // one-shot chats are prose: leave the model its own ceiling
      16384,
      options,
      false,
    ),
  })
  wd.touch()
  if (!response.ok) {
    return {
      ok: false,
      error: `HTTP ${response.status}: ${httpBodyDetail(await readCappedResponseText(response, () => wd.touch()))}`,
    }
  }
  const bodyText = await readCappedResponseText(response, () => wd.touch())
  let json: ResponsesObject
  try {
    json = JSON.parse(bodyText) as ResponsesObject
  } catch {
    return { ok: false, error: `AI returned a non-JSON response: ${httpBodyDetail(bodyText)}` }
  }
  if (json.error) return { ok: false, error: sseErrorText(json.error, 'Model error') }
  const content = (json.output ?? [])
    .filter((item) => item.type === 'message')
    .flatMap((item) => item.content ?? [])
    .filter((part) => part.type === 'output_text')
    .map((part) => part.text ?? '')
    .join('')
  if (!content) return { ok: false, error: 'AI returned an empty response' }
  return { ok: true, content }
}
