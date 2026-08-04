import type { AgentMessage, AgentToolCall, AgentToolDef } from '@genoffice/agent-core'
import { httpBodyDetail } from './http-error'
import { GENSPARK_LLM_BASE_URLS } from './providers'
import type { AiProviderConfig, AiProviderId } from './types'

// ---- streaming (SSE line splitting shared by all providers) ----

export async function* sseLines(
  body: NodeJS.ReadableStream | ReadableStream<Uint8Array>,
): AsyncGenerator<string> {
  const decoder = new TextDecoder()
  let buffer = ''
  const stream = body as ReadableStream<Uint8Array>
  const reader = stream.getReader()
  while (true) {
    const { done, value } = await reader.read()
    if (done) break
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

export async function streamAnthropic(
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
  baseUrl = 'https://api.anthropic.com',
): Promise<void> {
  let response: Response
  try {
    response = await fetch(`${baseUrl.replace(/\/$/, '')}/v1/messages`, {
      method: 'POST',
      signal: cb.signal,
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': config.apiKey,
        'anthropic-version': '2023-06-01',
        // Fetch in the Electron main process goes through Chromium's network stack, which adds
        // browser-semantics headers; Anthropic rejects those with 403 "Request not allowed". This
        // header is the official opt-in for direct access from browser/Electron environments.
        'anthropic-dangerous-direct-browser-access': 'true',
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
  if (!response.ok || !response.body) {
    throw new Error(`Claude HTTP ${response.status}: ${httpBodyDetail(await response.text())}`)
  }
  // tool_use inputs stream as partial JSON per content block
  const pendingTools = new Map<number, { id: string; name: string; json: string }>()
  for await (const line of sseLines(response.body)) {
    if (!line.startsWith('data:')) continue
    const payload = line.slice(5).trim()
    if (!payload) continue
    const event = JSON.parse(payload) as {
      type: string
      index?: number
      content_block?: { type?: string; id?: string; name?: string }
      delta?: { type?: string; text?: string; partial_json?: string }
      error?: { message?: string }
    }
    if (event.type === 'content_block_start' && event.content_block?.type === 'tool_use') {
      pendingTools.set(event.index ?? 0, {
        id: event.content_block.id ?? crypto.randomUUID(),
        name: event.content_block.name ?? '',
        json: '',
      })
    } else if (event.type === 'content_block_delta') {
      if (event.delta?.type === 'text_delta' && event.delta.text) cb.onDelta(event.delta.text)
      else if (event.delta?.type === 'input_json_delta') {
        const pending = pendingTools.get(event.index ?? 0)
        if (pending) pending.json += event.delta.partial_json ?? ''
      }
    } else if (event.type === 'content_block_stop') {
      const pending = pendingTools.get(event.index ?? 0)
      if (pending) {
        pendingTools.delete(event.index ?? 0)
        const { input, error } = parseToolInput(pending.json)
        cb.onToolCall({ id: pending.id, name: pending.name, input, inputError: error })
      }
    } else if (event.type === 'error') {
      throw new Error(event.error?.message ?? 'Claude stream error')
    }
  }
}

// ---- Gemini ----

/**
 * Gemini's function-declaration schemas are a proto-based OpenAPI subset, not
 * full JSON Schema. It rejects anything outside that subset — and one bad
 * declaration fails the *whole* request. Normalizations applied recursively:
 *   - `type` unions (`["string","null"]`) → first non-null scalar + `nullable`
 *   - `additionalProperties` → removed (unsupported)
 *   - `anyOf`/`oneOf`/`allOf` → collapsed to their first non-null member
 *   - tuple-style `items` (array) → first element (the proto field is singular)
 */
export function sanitizeGeminiSchema(node: unknown): unknown {
  if (Array.isArray(node)) {
    return node.map((n) => sanitizeGeminiSchema(n))
  }
  if (typeof node !== 'object' || node === null) return node
  const schema = { ...(node as Record<string, unknown>) }

  delete schema.additionalProperties
  delete schema.$schema
  delete schema.propertyNames
  delete schema.definitions

  if (Array.isArray(schema.type)) {
    const union = schema.type as string[]
    const hasNull = union.includes('null')
    const primary = union.find((t) => t !== 'null') ?? 'string'
    schema.type = primary
    if (hasNull) schema.nullable = true
  }

  for (const key of ['anyOf', 'oneOf', 'allOf'] as const) {
    const alternatives = schema[key]
    if (Array.isArray(alternatives)) {
      const pick = alternatives.find(
        (alt) => typeof alt === 'object' && alt !== null && !(alt as { type?: string }).type?.includes('null'),
      ) ?? alternatives[0]
      if (typeof pick === 'object' && pick !== null) {
        const merged = { ...schema, ...(pick as Record<string, unknown>) }
        delete merged[key]
        return sanitizeGeminiSchema(merged)
      }
      delete schema[key]
    }
  }

  if (Array.isArray(schema.items)) {
    schema.items = schema.items[0]
  }
  if (schema.items) {
    schema.items = sanitizeGeminiSchema(schema.items)
  }
  if (schema.properties && typeof schema.properties === 'object') {
    const props = schema.properties as Record<string, unknown>
    const sanitized: Record<string, unknown> = {}
    for (const [key, value] of Object.entries(props)) {
      sanitized[key] = sanitizeGeminiSchema(value)
    }
    schema.properties = sanitized
  }
  if (schema.enum && Array.isArray(schema.enum)) {
    schema.enum = schema.enum.map((v) => (typeof v === 'string' ? v : String(v)))
  }
  return schema
}

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

export async function streamGemini(
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
  baseUrl = 'https://generativelanguage.googleapis.com/v1beta',
): Promise<void> {
  const url = `${baseUrl.replace(/\/$/, '')}/models/${config.model}:streamGenerateContent?alt=sse`
  const response = await fetch(url, {
    method: 'POST',
    signal: cb.signal,
    headers: { 'Content-Type': 'application/json', 'x-goog-api-key': config.apiKey },
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
                  parameters: sanitizeGeminiSchema(t.inputSchema),
                })),
              },
            ],
          }
        : {}),
      generationConfig: { temperature: 0.3, maxOutputTokens: maxTokens },
    }),
  })
  if (!response.ok || !response.body) {
    throw new Error(`Gemini HTTP ${response.status}: ${httpBodyDetail(await response.text())}`)
  }
  for await (const line of sseLines(response.body)) {
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
      }>
    }
    for (const part of event.candidates?.[0]?.content?.parts ?? []) {
      if (part.text) cb.onDelta(part.text)
      // Gemini emits function calls whole, never as partial JSON
      if (part.functionCall?.name) {
        cb.onToolCall({
          id: crypto.randomUUID(),
          name: part.functionCall.name,
          input: part.functionCall.args ?? {},
        })
      }
    }
  }
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
      out.push({
        role: 'assistant',
        content: m.text || null,
        ...(m.toolCalls && m.toolCalls.length > 0
          ? {
              tool_calls: m.toolCalls.map((call) => ({
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

export async function streamOpenAiCompatible(
  baseUrl: string,
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
): Promise<void> {
  const response = await fetch(`${baseUrl.replace(/\/$/, '')}/chat/completions`, {
    method: 'POST',
    signal: cb.signal,
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${config.apiKey}`,
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
  if (!response.ok || !response.body) {
    throw new Error(`HTTP ${response.status}: ${httpBodyDetail(await response.text())}`)
  }
  // tool call arguments stream in fragments keyed by index
  const pendingTools = new Map<number, { id: string; name: string; json: string }>()
  const flushTools = () => {
    for (const pending of pendingTools.values()) {
      if (pending.name) {
        const { input, error } = parseToolInput(pending.json)
        cb.onToolCall({ id: pending.id, name: pending.name, input, inputError: error })
      }
    }
    pendingTools.clear()
  }
  for await (const line of sseLines(response.body)) {
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
    }
    const choice = event.choices?.[0]
    if (!choice) continue
    if (choice.delta?.content) cb.onDelta(choice.delta.content)
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
    if (choice.finish_reason) flushTools()
  }
  flushTools()
}

const OPENAI_COMPATIBLE_BASE_URLS: Partial<Record<AiProviderId, string>> = {
  deepseek: 'https://api.deepseek.com/v1',
  openai: 'https://api.openai.com/v1',
  openrouter: 'https://openrouter.ai/api/v1',
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
    case 'openrouter':
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
