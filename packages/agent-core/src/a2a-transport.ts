import type {
  AgentMessage,
  AgentStreamCallbacks,
  AgentStreamRequest,
  AgentTransport,
} from './types'

// ---- A2A protocol wire types (subset used by this transport) ----
// Spec: https://github.com/a2aproject/A2A

interface A2ATextPart {
  type: 'text'
  text: string
}

interface A2ADataPart {
  type: 'data'
  data: Record<string, unknown>
}

type A2APart = A2ATextPart | A2ADataPart

interface A2AMessage {
  role: 'user' | 'agent'
  parts: A2APart[]
}

interface A2ATaskStatus {
  state: 'submitted' | 'working' | 'input-required' | 'completed' | 'failed' | 'canceled'
  message?: A2AMessage
  timestamp?: string
}

interface A2AStreamEvent {
  id: string
  status?: A2ATaskStatus
  artifact?: {
    name: string
    parts: A2APart[]
    index?: number
    append?: boolean
    lastChunk?: boolean
  }
  final?: boolean
}

interface A2AJsonRpcResponse {
  jsonrpc: '2.0'
  id: unknown
  result?: A2AStreamEvent
  error?: { code: number; message: string; data?: unknown }
}

// ---- message encoding ----

/**
 * Encode GenOffice conversation history into a single A2A user message.
 * System prompt is prepended as a bracketed text part; assistant turns and
 * tool results are encoded as data parts so the remote agent has full context
 * without any special multi-turn session wiring.
 */
function encodeAsA2AMessage(system: string, messages: AgentMessage[]): A2AMessage {
  const parts: A2APart[] = []

  if (system) {
    parts.push({ type: 'text', text: `[System]\n${system}` })
  }

  for (const msg of messages) {
    if (msg.role === 'user') {
      parts.push({ type: 'text', text: msg.text })
    } else if (msg.role === 'assistant') {
      parts.push({
        type: 'data',
        data: {
          role: 'assistant',
          text: msg.text,
          ...(msg.toolCalls?.length ? { toolCalls: msg.toolCalls } : {}),
        },
      })
    } else {
      // tool results
      parts.push({ type: 'data', data: { role: 'tool', results: msg.results } })
    }
  }

  return { role: 'user', parts }
}

// ---- SSE parsing ----

async function* sseEvents(
  body: ReadableStream<Uint8Array>,
  signal: AbortSignal,
): AsyncGenerator<string> {
  const reader = body.getReader()
  const decoder = new TextDecoder()
  let buf = ''
  try {
    while (!signal.aborted) {
      const { value, done } = await reader.read()
      if (done) break
      buf += decoder.decode(value, { stream: true })
      const lines = buf.split('\n')
      buf = lines.pop() ?? ''
      for (const line of lines) {
        if (line.startsWith('data:')) yield line.slice(5).trim()
      }
    }
  } finally {
    reader.releaseLock()
  }
  if (buf.startsWith('data:')) yield buf.slice(5).trim()
}

// ---- public API ----

export interface A2ATransportOptions {
  /**
   * Base URL of the A2A agent (e.g. "http://localhost:9100").
   * The transport POSTs JSON-RPC to this URL directly.
   */
  agentUrl: string
  /** Bearer token sent as Authorization header when provided. */
  apiKey?: string | undefined
  /**
   * Session ID threaded through all tasks so the remote agent can correlate
   * turns. Callers should create one UUID per AgentLoop instance.
   */
  sessionId?: string | undefined
  /** Extra request headers (e.g. X-Agent-Type attribution). */
  headers?: Record<string, string> | undefined
}

/**
 * AgentTransport that talks to a remote A2A-protocol agent over HTTP SSE.
 *
 * The full GenOffice conversation context (system prompt + message history)
 * is encoded into a single A2A user message on every transport.stream() call.
 * The remote agent handles its own model calls and tools; GenOffice receives
 * streamed text artifacts back as onDelta events.
 *
 * Spec: https://github.com/a2aproject/A2A
 */
export function createA2ATransport(options: A2ATransportOptions): AgentTransport {
  const { agentUrl, apiKey, sessionId, headers: extraHeaders = {} } = options
  const baseUrl = agentUrl.replace(/\/$/, '')

  const authHeaders: Record<string, string> = apiKey ? { Authorization: `Bearer ${apiKey}` } : {}

  return {
    stream(request: AgentStreamRequest, cb: AgentStreamCallbacks) {
      const controller = new AbortController()
      const taskId = crypto.randomUUID()

      const body = JSON.stringify({
        jsonrpc: '2.0',
        id: taskId,
        method: 'tasks/sendSubscribe',
        params: {
          id: taskId,
          ...(sessionId ? { sessionId } : {}),
          message: encodeAsA2AMessage(request.system, request.messages),
        },
      })

      void runStream(controller.signal, body, cb)

      return { cancel: () => controller.abort() }
    },
  }

  async function runStream(
    signal: AbortSignal,
    body: string,
    cb: AgentStreamCallbacks,
  ): Promise<void> {
    let resp: Response
    try {
      resp = await fetch(baseUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Accept: 'text/event-stream',
          ...authHeaders,
          ...extraHeaders,
        },
        body,
        signal,
      })
    } catch (err) {
      if (!signal.aborted) {
        cb.onError(err instanceof Error ? err.message : 'A2A fetch failed')
      }
      return
    }

    if (!resp.ok) {
      cb.onError(`A2A agent returned HTTP ${resp.status}`)
      return
    }

    if (!resp.body) {
      cb.onError('A2A agent returned no response body')
      return
    }

    let settled = false
    const settle = () => {
      settled = true
    }

    for await (const data of sseEvents(resp.body, signal)) {
      if (settled || !data || data === '[DONE]') continue

      let parsed: A2AJsonRpcResponse
      try {
        parsed = JSON.parse(data) as A2AJsonRpcResponse
      } catch {
        continue
      }

      if (parsed.error) {
        settle()
        cb.onError(parsed.error.message)
        return
      }

      const ev = parsed.result
      if (!ev) continue

      // Status events
      if (ev.status) {
        const { state, message } = ev.status
        if (state === 'failed' || state === 'canceled') {
          settle()
          const errText = message?.parts.find((p): p is A2ATextPart => p.type === 'text')?.text
          cb.onError(errText ?? `A2A task ${state}`)
          return
        }
        if (state === 'completed' && !ev.artifact) {
          settle()
          cb.onDone()
          return
        }
      }

      // Artifact events (streamed text output)
      if (ev.artifact) {
        for (const part of ev.artifact.parts) {
          if (part.type === 'text' && part.text) {
            cb.onDelta(part.text)
          }
        }
        if (ev.artifact.lastChunk || ev.final) {
          settle()
          cb.onDone()
          return
        }
      }

      if (ev.final) {
        settle()
        cb.onDone()
        return
      }
    }

    if (!settled && !signal.aborted) {
      cb.onDone()
    }
  }
}
