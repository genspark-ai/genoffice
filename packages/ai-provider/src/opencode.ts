/**
 * OpenCode provider: drives a local `opencode serve` process over its native
 * v2 HTTP API (the only API the CLI exposes — there is no OpenAI-compatible
 * endpoint). The agent loop runs entirely inside opencode: we hand it the
 * GenOffice system prompt plus the conversation transcript as plain text, and
 * stream its final text back. Every session is created with permission
 * deny-all, so opencode's own tools (bash, file edits…) can never touch the
 * machine from inside an office workflow.
 */

import type { ChildProcess } from 'node:child_process'
import type { AgentMessage, AgentToolCall, AgentToolDef } from '@genoffice/agent-core'
import { aiFetch } from './fetch'
import { httpBodyDetail } from './http-error'
import type { AiChatResponse, AiProviderConfig } from './types'
import {
  AI_CHAT_RESPONSE_TIMEOUT_MS,
  createStreamWatchdog,
  type StreamWatchdog,
} from './watchdog'

export const OPENCODE_DEFAULT_BASE_URL = 'http://127.0.0.1:3456'
const OPENCODE_SERVER_START_TIMEOUT_MS = 20_000
const OPENCODE_SERVER_POLL_MS = 400

export class OpencodeServerError extends Error {
  constructor(message: string, options?: { cause?: unknown }) {
    super(message, options)
    this.name = 'OpencodeServerError'
  }
}

// ---- server lifecycle ------------------------------------------------------

let spawnedServer: { child: ChildProcess; baseUrl: string } | null = null

function portFromBaseUrl(baseUrl: string): number {
  try {
    const { port } = new URL(baseUrl)
    return port ? Number(port) : 3456
  } catch {
    return 3456
  }
}

/** Is an opencode server answering on this base URL? (health probe) */
async function serverReachable(baseUrl: string): Promise<boolean> {
  try {
    const res = await aiFetch(`${baseUrl.replace(/\/$/, '')}/api/session?limit=1`, {
      signal: AbortSignal.timeout(2_000),
    })
    return res.ok
  } catch {
    return false
  }
}

function killChild(child: ChildProcess): void {
  try {
    if (!child.killed) child.kill()
  } catch {
    /* already gone */
  }
}

/**
 * Make sure an opencode server is reachable at `baseUrl`; when none is, spawn
 * `opencode serve` (the CLI must be on PATH) and wait for it to answer.
 * Concurrent callers share one spawn; the child is killed on process exit.
 */
export async function ensureOpencodeServer(
  baseUrl = OPENCODE_DEFAULT_BASE_URL,
): Promise<void> {
  const normalized = baseUrl.replace(/\/$/, '')
  if (await serverReachable(normalized)) return
  if (spawnedServer?.baseUrl === normalized) return

  const port = portFromBaseUrl(normalized)
  const cli = process.platform === 'win32' ? 'opencode.cmd' : 'opencode'
  let child: ChildProcess
  try {
    // Dynamic import keeps node:child_process out of the renderer bundle: the
    // spawn path only ever runs in main / preload / tests.
    const { spawn } = await import('node:child_process')
    child = spawn(cli, ['serve', '--port', String(port), '--hostname', '127.0.0.1'], {
      stdio: 'ignore',
    })
  } catch (err) {
    throw new OpencodeServerError(
      `Could not start the opencode CLI (${cli}): ${String(err)}`,
      { cause: err },
    )
  }
  spawnedServer = { child, baseUrl: normalized }

  let spawnError: Error | null = null
  child.once('error', (err) => {
    spawnError = new OpencodeServerError(
      `Could not start the opencode CLI (${cli}). Install opencode or start "opencode serve" yourself, then retry. (${err.message})`,
      { cause: err },
    )
  })

  const deadline = Date.now() + OPENCODE_SERVER_START_TIMEOUT_MS
  while (Date.now() < deadline) {
    if (spawnError) throw spawnError
    if (await serverReachable(normalized)) return
    await new Promise((r) => setTimeout(r, OPENCODE_SERVER_POLL_MS))
  }
  killChild(child)
  spawnedServer = null
  throw new OpencodeServerError(
    `The opencode server at ${normalized} did not become ready within ${Math.round(OPENCODE_SERVER_START_TIMEOUT_MS / 1000)}s.`,
  )
}

/** Kill a server this module spawned (used by tests and app teardown). */
export function disposeOpencodeServer(): void {
  if (spawnedServer) {
    killChild(spawnedServer.child)
    spawnedServer = null
  }
}

if (typeof process !== 'undefined') {
  process.on('exit', () => {
    if (spawnedServer) killChild(spawnedServer.child)
  })
}

// ---- prompt text -----------------------------------------------------------

/** "opencode/model-id" or "model-id" → { id, providerID }; empty → null (server default) */
export function parseOpencodeModel(spec: string): { id: string; providerID: string } | null {
  const s = spec.trim()
  if (!s) return null
  const idx = s.indexOf('/')
  if (idx === -1) return { id: s, providerID: 'opencode' }
  return { id: s.slice(idx + 1), providerID: s.slice(0, idx) }
}

/**
 * Serialize a GenOffice turn into the plain-text prompt handed to the opencode
 * agent: the system prompt (document context + editing instructions) followed
 * by the full conversation transcript. opencode runs its own agent loop, so
 * tool results from earlier GenOffice turns are included as transcript text.
 */
export function opencodePromptText(system: string, messages: AgentMessage[]): string {
  const parts: string[] = []
  if (system.trim()) parts.push(system.trim())
  if (messages.length > 0) {
    const transcript: string[] = []
    for (const m of messages) {
      if (m.role === 'user') {
        transcript.push(`## User\n${m.text}`)
      } else if (m.role === 'assistant') {
        const calls = (m.toolCalls ?? [])
          .map((t) => `- ${t.name}(${JSON.stringify(t.input)})`)
          .join('\n')
        transcript.push(
          `## Assistant\n${m.text ?? ''}${calls ? `\n\nTool calls made:\n${calls}` : ''}`,
        )
      } else {
        transcript.push(
          `## Tool results\n${m.results
            .map((r) => `- ${r.name}${r.isError ? ' (error)' : ''}: ${r.output}`)
            .join('\n')}`,
        )
      }
    }
    parts.push(`# Conversation history\n\n${transcript.join('\n\n')}`)
    parts.push(
      '# Task\nContinue as the assistant. Do not repeat what has already been done or stated — produce the next assistant message or the final response.',
    )
  }
  return parts.join('\n\n')
}

// ---- SSE plumbing (local copy: avoids a circular import with stream.ts) ----

async function* sseDataLines(
  body: ReadableStream<Uint8Array>,
  onBytes: () => void,
): AsyncGenerator<string> {
  const decoder = new TextDecoder()
  let buffer = ''
  const reader = body.getReader()
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    onBytes()
    buffer += decoder.decode(value, { stream: true })
    const lines = buffer.split('\n')
    buffer = lines.pop() ?? ''
    for (const line of lines) {
      const trimmed = line.trim()
      if (trimmed.startsWith('data:')) yield trimmed.slice(5).trim()
    }
  }
  if (buffer.trim().startsWith('data:')) yield buffer.trim().slice(5).trim()
}

interface OpencodeEvent {
  type?: string
  text?: string
  finish?: string
  error?: unknown
  data?: { text?: string; finish?: string; error?: unknown; message?: string }
}

function sseErrorText(value: unknown, fallback: string): string {
  if (typeof value === 'string' && value) return value
  if (value && typeof value === 'object') {
    const message = (value as { message?: unknown }).message
    if (typeof message === 'string' && message) return message
    try {
      return JSON.stringify(value)
    } catch {
      /* fall through */
    }
  }
  return fallback
}

// ---- model resolution ------------------------------------------------------

/**
 * Pick the model for a session when the user left the model field blank. An
 * `opencode serve` without a configured default model accepts the prompt and
 * then never produces output, which surfaces to the user as a network idle
 * timeout. Ask the server for its config: if it has a default model, keep the
 * model-less create (the server default wins); otherwise fall back to the
 * first model it advertises on /api/model. Any failure → null → the caller
 * creates the session without a model, preserving old behaviour for servers
 * that predate the /config and /api/model endpoints.
 */
async function resolveOpencodeModel(
  baseUrl: string,
  signal: AbortSignal,
): Promise<{ id: string; providerID: string } | null> {
  try {
    const configRes = await aiFetch(`${baseUrl}/config`, { signal })
    if (configRes.ok) {
      const cfg = (await configRes.json()) as { model?: unknown }
      if (typeof cfg.model === 'string' && cfg.model.trim()) return null
    }
  } catch {
    // config endpoint unreachable — fall through to the model list
  }
  try {
    const modelRes = await aiFetch(`${baseUrl}/api/model`, { signal })
    if (!modelRes.ok) return null
    const body = (await modelRes.json()) as {
      data?: Array<{ id?: string; providerID?: string }>
    }
    const first = body.data?.find((m) => m.id)
    return first?.id ? { id: first.id, providerID: first.providerID ?? 'opencode' } : null
  } catch {
    return null
  }
}

// ---- the turn --------------------------------------------------------------

async function opencodeTurn(
  baseUrl: string,
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  cb: {
    onDelta: (text: string) => void
    onActivity?: () => void
    signal: AbortSignal
  },
): Promise<string> {
  const normalized = baseUrl.replace(/\/$/, '')
  await ensureOpencodeServer(normalized)

  let model = parseOpencodeModel(config.model)
  if (!model) model = await resolveOpencodeModel(normalized, cb.signal)
  const createResponse = await aiFetch(`${normalized}/api/session`, {
    method: 'POST',
    signal: cb.signal,
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      title: 'GenOffice',
      ...(model ? { model } : {}),
      // deny everything: the opencode agent must never run tools from inside an office doc
      permission: [{ permission: '*', action: 'deny', pattern: '*' }],
    }),
  })
  cb.onActivity?.()
  if (!createResponse.ok) {
    throw new OpencodeServerError(
      `opencode: failed to create a session (HTTP ${createResponse.status}): ${httpBodyDetail(await createResponse.text())}`,
    )
  }
  const session = (await createResponse.json()) as {
    sessionID?: string
    data?: { id?: string }
  }
  // opencode 1.18+ wraps the id in a data object; older builds returned it
  // top-level. Accept both so the provider survives either server version.
  const sessionId = session.sessionID ?? session.data?.id
  if (!sessionId) {
    throw new OpencodeServerError('opencode: session create response had no sessionID')
  }

  // best-effort: tell the server to stop when the caller aborts; the fetch below
  // also aborts on the same signal, which is what actually unwinds the read
  const interrupt = () => {
    void aiFetch(`${normalized}/api/session/${sessionId}/interrupt`, { method: 'POST' }).catch(
      () => {},
    )
  }
  if (cb.signal.aborted) interrupt()
  else cb.signal.addEventListener('abort', interrupt, { once: true })

  try {
    const promptResponse = await aiFetch(`${normalized}/api/session/${sessionId}/prompt`, {
      method: 'POST',
      signal: cb.signal,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ prompt: { text: opencodePromptText(system, messages) } }),
    })
    cb.onActivity?.()
    if (!promptResponse.ok) {
      throw new OpencodeServerError(
        `opencode: the prompt was rejected (HTTP ${promptResponse.status}): ${httpBodyDetail(await promptResponse.text())}`,
      )
    }

    const eventResponse = await aiFetch(`${normalized}/api/session/${sessionId}/event`, {
      method: 'GET',
      signal: cb.signal,
    })
    cb.onActivity?.()
    if (!eventResponse.ok || !eventResponse.body) {
      throw new OpencodeServerError(
        `opencode: the event stream failed (HTTP ${eventResponse.status}): ${httpBodyDetail(await eventResponse.text())}`,
      )
    }

    let output = ''
    let sawStop = false
    let sawStep = false
    for await (const payload of sseDataLines(eventResponse.body, () => cb.onActivity?.())) {
      if (!payload) continue
      let event: OpencodeEvent
      try {
        event = JSON.parse(payload) as OpencodeEvent
      } catch {
        continue // non-JSON keepalive frame
      }
      switch (event.type) {
        case 'session.next.text.ended':
          // 1.18+ puts the payload in a nested data object; older builds had it top-level
          {
            const text = event.text ?? event.data?.text
            if (text) {
              output += text
              cb.onDelta(text)
            }
          }
          break
        case 'session.next.step.ended':
          sawStep = true
          if ((event.finish ?? event.data?.finish) === 'stop') sawStop = true
          if (event.error ?? event.data?.error ?? event.data?.message) {
            throw new OpencodeServerError(
              `opencode: ${sseErrorText(
                event.error ?? event.data?.error ?? event.data?.message,
                'agent error',
              )}`,
            )
          }
          break
        case 'session.next.error':
          throw new OpencodeServerError(
            `opencode: ${sseErrorText(
              event.error ?? event.data?.error ?? event.data?.message,
              'agent error',
            )}`,
          )
        default:
          // everything else (step.started, tool.called with executed:false, …) is
          // loop bookkeeping — the text.ended / step.ended events carry the result
          break
      }
      if (sawStop) break
    }

    // stream ended (or stopped) without any step framing → treat as empty, not as a
    // success: a genuine run always terminates with a step.ended
    if (!sawStep && !sawStop && !output) {
      throw new OpencodeServerError('opencode returned no content (empty stream)')
    }
    return output
  } finally {
    cb.signal.removeEventListener('abort', interrupt)
  }
}

/** Stream one agent turn from a local opencode server. */
export async function streamOpencode(
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  _tools: AgentToolDef[], // opencode runs its own agent loop; GenOffice tools are not forwarded
  _maxTokens: number,
  cb: {
    onDelta: (text: string) => void
    onToolCall: (call: AgentToolCall) => void
    onStopReason?: (reason: string) => void
    onActivity?: () => void
    signal: AbortSignal
  },
): Promise<void> {
  const wd = createStreamWatchdog(cb.signal)
  await wd.guard(async () => {
    const baseUrl = config.baseUrl || OPENCODE_DEFAULT_BASE_URL
    await opencodeTurn(baseUrl, config, system, messages, {
      onDelta: cb.onDelta,
      ...(cb.onActivity ? { onActivity: cb.onActivity } : {}),
      signal: wd.signal,
    })
  })
  cb.onStopReason?.('stop')
}

/** One-shot (non-streaming) chat call against a local opencode server. */
export async function chatOpencode(
  config: AiProviderConfig,
  system: string,
  user: string,
  signal?: AbortSignal,
): Promise<AiChatResponse> {
  const wd = createStreamWatchdog(signal, AI_CHAT_RESPONSE_TIMEOUT_MS)
  return wd.guard(async () => {
    const baseUrl = config.baseUrl || OPENCODE_DEFAULT_BASE_URL
    try {
      const content = await opencodeTurn(
        baseUrl,
        config,
        system,
        [{ role: 'user', text: user }],
        {
          onDelta: () => {},
          signal: wd.signal,
        },
      )
      if (!content) return { ok: false, error: 'opencode returned an empty response' }
      return { ok: true, content }
    } catch (err) {
      return { ok: false, error: err instanceof Error ? err.message : String(err) }
    }
  })
}
