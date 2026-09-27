import type { AgentToolCall } from '@genoffice/agent-core'

// ---- endpoint URL composition ----

/**
 * Append a provider endpoint path to a base URL without string concatenation.
 *
 * A custom base URL may carry a query string (Azure-style `?api-version=…`,
 * gateways that pin a version) and a fragment. Concatenating would push the
 * path into the query component, so the request 404s or lands on the wrong
 * route. Setting `pathname` keeps the query in place and the fragment is
 * dropped — it is never sent to the server anyway.
 */
export function endpointUrl(baseUrl: string, path: string, search?: string): string {
  const url = new URL(baseUrl)
  const base = url.pathname.replace(/\/+$/, '')
  const suffix = path.replace(/^\/+/, '')
  url.pathname = suffix ? `${base}/${suffix}` : base
  url.hash = ''
  if (search !== undefined) url.search = search
  return url.toString()
}

// ---- streaming (SSE line splitting shared by all providers) ----

/** Max buffered SSE line: a gateway sending GB without newline would OOM main. */
export const MAX_SSE_LINE_BYTES = 4 * 1024 * 1024
export const MAX_RESPONSE_BODY_BYTES = MAX_SSE_LINE_BYTES

export class ResponseBodyTooLargeError extends Error {
  constructor(
    readonly receivedBytes: number,
    readonly capBytes: number,
  ) {
    super(`Response body exceeded the ${capBytes}-byte limit (${receivedBytes} bytes)`)
    this.name = 'ResponseBodyTooLargeError'
  }
}

export async function readCappedResponseText(
  response: Response,
  onBytes?: () => void,
): Promise<string> {
  const declaredBytes = Number(response.headers.get('content-length'))
  if (Number.isFinite(declaredBytes) && declaredBytes > MAX_RESPONSE_BODY_BYTES) {
    if (response.body) await response.body.cancel().catch(() => undefined)
    throw new ResponseBodyTooLargeError(declaredBytes, MAX_RESPONSE_BODY_BYTES)
  }
  if (!response.body) return ''

  const reader = response.body.getReader()
  const decoder = new TextDecoder()
  let text = ''
  let bytes = 0
  try {
    for (;;) {
      const { done, value } = await reader.read()
      if (done) break
      onBytes?.()
      bytes += value.byteLength
      if (bytes > MAX_RESPONSE_BODY_BYTES) {
        throw new ResponseBodyTooLargeError(bytes, MAX_RESPONSE_BODY_BYTES)
      }
      text += decoder.decode(value, { stream: true })
    }
    return text + decoder.decode()
  } finally {
    await reader.cancel().catch(() => undefined)
    reader.releaseLock()
  }
}

export async function* sseLines(
  body: NodeJS.ReadableStream | ReadableStream<Uint8Array>,
  onBytes?: () => void,
): AsyncGenerator<string> {
  const decoder = new TextDecoder()
  let buffer = ''
  let lineBytes = 0
  const stream = body as ReadableStream<Uint8Array>
  const reader = stream.getReader()
  try {
    while (true) {
      const { done, value } = await reader.read()
      if (done) break
      onBytes?.()
      for (const byte of value) {
        if (byte === 0x0a) {
          lineBytes = 0
          continue
        }
        lineBytes++
        if (lineBytes > MAX_SSE_LINE_BYTES) {
          throw new Error(
            `SSE line exceeded buffer limit (${lineBytes} bytes, cap ${MAX_SSE_LINE_BYTES}); the gateway sent a line without newline.`,
          )
        }
      }
      buffer += decoder.decode(value, { stream: true })
      const lines = buffer.split('\n')
      buffer = lines.pop() ?? ''
      for (const line of lines) yield line
    }
    // Flush the decoder: bytes of a multibyte char still buffered inside
    // TextDecoder under { stream: true } are discarded without a final
    // decode() — a stream cut mid-char (dropped connection) would lose its
    // tail silently instead of surfacing the standard replacement mark.
    buffer += decoder.decode()
    if (buffer) yield buffer
  } finally {
    // The consumer may abandon this generator mid-stream (an in-band gateway
    // error thrown inside the for-await loop calls .return()). Without this
    // cleanup the reader stays locked and the underlying socket is not
    // returned to the pool until GC nondeterministically finalizes it.
    await reader.cancel().catch(() => undefined)
    reader.releaseLock()
  }
}

/**
 * Per-tool streamed argument buffer cap: a provider streaming argument
 * fragments forever (never finishing) would otherwise grow the pending
 * tool-call buffer without bound. Throwing aborts the turn; sseLines
 * cancels the underlying stream on the way out.
 */
export const MAX_TOOL_JSON_CHARS = 512_000

export function throwIfToolJsonOverBudget(jsonLength: number, provider: string): void {
  if (jsonLength > MAX_TOOL_JSON_CHARS) {
    throw new Error(
      `Tool call arguments exceeded the ${provider} buffer limit (${jsonLength} chars, cap ${MAX_TOOL_JSON_CHARS}); ` +
        'the provider kept streaming argument fragments without finishing. Ask for the output in several smaller parts.',
    )
  }
}

/**
 * Max tool calls started per streamed turn: per-argument bytes are capped
 * above, but a gateway could still stream 100k near-empty tool_use blocks
 * and grow the completed-call list without bound.
 */
export const MAX_STREAM_TOOL_CALLS = 100

export function throwIfToolCountOverBudget(count: number, provider: string): void {
  if (count > MAX_STREAM_TOOL_CALLS) {
    throw new Error(
      `Too many streamed tool calls (${count}, cap ${MAX_STREAM_TOOL_CALLS}) for ${provider}; ` +
        'the provider kept starting tool calls without finishing the turn.',
    )
  }
}

export interface StreamCallbacks {
  onDelta: (text: string) => void
  onToolCall: (call: AgentToolCall) => void
  /** raw model reasoning deltas (reasoning_content); stored so interleaved-thinking models get it echoed back */
  onReasoningDelta?: (text: string) => void
  /** normalized stop reason ('max_tokens' when the output was cut off by the token limit) */
  onStopReason?: (reason: string) => void
  /** bytes arrived on the wire (fires per network chunk, including SSE pings; used for keepalive) */
  onActivity?: () => void
  /** Stable renderer transport id for providers with native sessions. */
  sessionId?: string
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
export function sseErrorText(error: unknown, fallback: string): string {
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
export async function jsonBodyInsteadOfSse(
  response: Response,
  onBytes?: () => void,
): Promise<string | null> {
  const contentType = response.headers.get('content-type') ?? ''
  return contentType.toLowerCase().includes('application/json')
    ? readCappedResponseText(response, onBytes)
    : null
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

export function throwIfCreditsNotice(bodyText: string): void {
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
export function parseToolInput(json: string): { input: Record<string, unknown>; error?: string } {
  if (!json.trim()) return { input: {} }
  let parsed: unknown
  try {
    parsed = JSON.parse(json)
  } catch (e) {
    try {
      parsed = JSON.parse(repairUnescapedQuotes(json))
    } catch {
      const msg = e instanceof Error ? e.message : String(e)
      return { input: {}, error: `${msg}; raw: ${json.slice(0, 500)}` }
    }
  }
  if (!isPlainObject(parsed)) {
    return { input: {}, error: `tool input must be a JSON object; raw: ${json.slice(0, 500)}` }
  }
  return { input: parsed }
}

export function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
