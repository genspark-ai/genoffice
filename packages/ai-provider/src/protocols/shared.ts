import type { AgentToolCall } from '@genoffice/agent-core'

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
export async function jsonBodyInsteadOfSse(response: Response): Promise<string | null> {
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
