import type { AgentMessage, AgentToolDef } from '@genoffice/agent-core'
import { sseLines } from './stream'
import type {
  CodexAdapterRequest,
  CodexAuthContext,
  CodexCapabilities,
  CodexReasoningEffort,
} from './types'

export const CODEX_RESPONSES_URL = 'https://chatgpt.com/backend-api/codex/responses'
export const CODEX_MODELS_URL = 'https://chatgpt.com/backend-api/codex/models'
/** Codex wire-client version used by backend model eligibility filtering. */
export const CODEX_CLIENT_VERSION = '0.144.1'

const CODEX_REASONING_EFFORTS = new Set<CodexReasoningEffort>([
  'none',
  'minimal',
  'low',
  'medium',
  'high',
  'xhigh',
  'max',
  'ultra',
])

/** A status error with a deliberately minimal, safe 400 diagnostic. */
export class CodexHttpError extends Error {
  readonly status: number
  readonly code?: string
  readonly detail?: string

  constructor(status: number, body: string) {
    super(`Codex HTTP ${status}`)
    this.name = 'CodexHttpError'
    this.status = status
    if (status === 400) {
      const diagnostic = codex400Diagnostic(body)
      if (diagnostic.code) this.code = diagnostic.code
      if (diagnostic.detail) this.detail = diagnostic.detail
    }
  }
}

function codex400Diagnostic(body: string): { code?: string; detail?: string } {
  try {
    const parsed: unknown = JSON.parse(body)
    if (!parsed || typeof parsed !== 'object') return {}
    const root = parsed as Record<string, unknown>
    const error =
      root.error && typeof root.error === 'object' ? (root.error as Record<string, unknown>) : root
    const code = safeCode(error.code)
    const detail = safeDetail(error.message ?? error.detail)
    return { ...(code ? { code } : {}), ...(detail ? { detail } : {}) }
  } catch {
    return {}
  }
}

function safeCode(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined
  const code = value.trim()
  return /^[a-z][a-z0-9_.-]{0,63}$/i.test(code) ? code : undefined
}

function safeDetail(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined
  const detail = value.replace(/\s+/g, ' ').trim()
  // Refuse, rather than redact, a provider detail containing any secret or URL.
  if (
    !detail ||
    /https?:\/\/|\b(?:authorization|bearer|token|api[_ -]?key|password|secret|credential)\b|\b(?:sk-|sess-|eyJ)[a-z0-9_-]{8,}/i.test(
      detail,
    )
  ) {
    return undefined
  }
  return detail.slice(0, 160)
}

type CodexInput =
  | { role: 'user' | 'assistant'; content: string | Array<Record<string, string>> }
  | { type: 'function_call'; call_id: string; name: string; arguments: string }
  | { type: 'function_call_output'; call_id: string; output: string }

function codexInput(messages: AgentMessage[]): CodexInput[] {
  const input: CodexInput[] = []
  for (const message of messages) {
    if (message.role === 'user') {
      input.push({
        role: 'user',
        content: message.images?.length
          ? [
              ...(message.text ? [{ type: 'input_text', text: message.text }] : []),
              ...message.images.map((image) => ({
                type: 'input_image',
                image_url: `data:${image.mime};base64,${image.base64}`,
              })),
            ]
          : message.text,
      })
      continue
    }
    if (message.role === 'assistant') {
      if (message.text) {
        input.push({ role: 'assistant', content: [{ type: 'output_text', text: message.text }] })
      }
      for (const call of message.toolCalls ?? []) {
        input.push({
          type: 'function_call',
          call_id: call.id,
          name: call.name,
          arguments: JSON.stringify(call.input),
        })
      }
      continue
    }
    for (const result of message.results) {
      input.push({
        type: 'function_call_output',
        call_id: result.id,
        output: result.isError ? `Error: ${result.output}` : result.output,
      })
    }
  }
  return input
}

function codexTools(tools: AgentToolDef[]) {
  return tools.map((tool) => ({
    type: 'function' as const,
    name: tool.name,
    description: tool.description,
    parameters: tool.inputSchema,
    strict: false,
  }))
}

/** Build a stateless Codex Responses request. This layer never executes tools. */
export function buildCodexRequest(request: CodexAdapterRequest) {
  return {
    model: request.model,
    instructions: request.instructions,
    input: codexInput(request.messages),
    ...(request.tools.length > 0 ? { tools: codexTools(request.tools) } : {}),
    ...(request.reasoningEffort && request.reasoningEffort !== 'none'
      ? { reasoning: { effort: request.reasoningEffort } }
      : {}),
    store: false,
    stream: true,
  }
}

/** Fetch and validate the authenticated account's picker-safe model catalog. */
export async function fetchCodexCapabilities(
  auth: CodexAuthContext,
  clientVersion: string = CODEX_CLIENT_VERSION,
  fetchImpl: typeof fetch = fetch,
): Promise<CodexCapabilities> {
  const response = await fetchImpl(
    `${CODEX_MODELS_URL}?client_version=${encodeURIComponent(clientVersion)}`,
    {
      method: 'GET',
      headers: codexModelsHeaders(auth),
    },
  )
  if (!response.ok) throw new CodexHttpError(response.status, await response.text())
  const parsed: unknown = await response.json()
  const capabilities = parseCodexCapabilities(parsed)
  if (capabilities.models.length === 0)
    throw new Error('Codex models response has no selectable models')
  return capabilities
}

function parseCodexCapabilities(value: unknown): CodexCapabilities {
  if (
    !value ||
    typeof value !== 'object' ||
    !Array.isArray((value as { models?: unknown }).models)
  ) {
    throw new Error('Codex models response malformed')
  }
  const seen = new Set<string>()
  const models = [] as CodexCapabilities['models']
  for (const item of (value as { models: unknown[] }).models) {
    if (!item || typeof item !== 'object') continue
    const record = item as Record<string, unknown>
    const id = typeof record.slug === 'string' ? record.slug.trim() : ''
    if (
      !id ||
      id.length > 128 ||
      !/^[a-zA-Z0-9._-]+$/.test(id) ||
      record.visibility !== 'list' ||
      record.supported_in_api !== true ||
      seen.has(id)
    ) {
      continue
    }
    const reasoningEfforts = Array.isArray(record.supported_reasoning_levels)
      ? record.supported_reasoning_levels
          .map((preset) =>
            preset &&
            typeof preset === 'object' &&
            typeof (preset as { effort?: unknown }).effort === 'string'
              ? (preset as { effort: string }).effort
              : undefined,
          )
          .filter(
            (effort): effort is CodexReasoningEffort =>
              typeof effort === 'string' &&
              CODEX_REASONING_EFFORTS.has(effort as CodexReasoningEffort),
          )
      : []
    seen.add(id)
    models.push({ id, reasoningEfforts: [...new Set(reasoningEfforts)] })
  }
  return { models }
}

/** Headers required by Codex Responses; credentials remain in the main process. */
export function codexHeaders(auth: CodexAuthContext, sessionId: string): Record<string, string> {
  return {
    Accept: 'text/event-stream',
    Authorization: `Bearer ${auth.accessToken}`,
    'Content-Type': 'application/json',
    'OpenAI-Beta': 'responses=experimental',
    'chatgpt-account-id': auth.accountId,
    originator: 'codex_cli_rs',
    session_id: sessionId,
  }
}

function codexModelsHeaders(auth: CodexAuthContext): Record<string, string> {
  return {
    Accept: 'application/json',
    Authorization: `Bearer ${auth.accessToken}`,
    'chatgpt-account-id': auth.accountId,
    originator: 'codex_cli_rs',
  }
}

function abortError(): DOMException {
  return new DOMException('Codex request cancelled', 'AbortError')
}

function toolInput(argumentsJson: string): Record<string, unknown> {
  try {
    const parsed = JSON.parse(argumentsJson)
    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
      throw new Error('arguments must be a JSON object')
    }
    return parsed as Record<string, unknown>
  } catch {
    throw new Error(`Codex tool arguments malformed: ${argumentsJson.slice(0, 500)}`)
  }
}

/** Stream Codex Responses events into existing GenOffice callbacks; never execute tools. */
export async function streamCodexResponse(request: CodexAdapterRequest): Promise<void> {
  if (request.signal.aborted) throw abortError()
  const response = await fetch(CODEX_RESPONSES_URL, {
    method: 'POST',
    signal: request.signal,
    headers: codexHeaders(request.auth, crypto.randomUUID()),
    body: JSON.stringify(buildCodexRequest(request)),
  })
  if (!response.ok || !response.body) {
    throw new CodexHttpError(response.status, await response.text())
  }

  const pending = new Map<string, { name: string; arguments: string }>()
  for await (const line of sseLines(response.body)) {
    if (request.signal.aborted) throw abortError()
    if (!line.startsWith('data:')) continue
    const payload = line.slice(5).trim()
    if (!payload || payload === '[DONE]') continue
    let event: {
      type?: string
      delta?: string
      call_id?: string
      name?: string
      arguments?: string
      item?: { type?: string; call_id?: string; name?: string; arguments?: string }
      error?: { message?: string }
      response?: { error?: { message?: string } }
    }
    try {
      event = JSON.parse(payload) as typeof event
    } catch {
      throw new Error(`Codex stream malformed event: ${payload.slice(0, 500)}`)
    }
    if (event.type === 'response.output_text.delta' && event.delta) {
      request.onDelta(event.delta)
      continue
    }
    if (event.type === 'response.output_item.added' && event.item?.type === 'function_call') {
      const callId = event.item.call_id
      if (callId)
        pending.set(callId, { name: event.item.name ?? '', arguments: event.item.arguments ?? '' })
      continue
    }
    if (event.type === 'response.function_call_arguments.delta' && event.call_id) {
      const call = pending.get(event.call_id)
      if (call) call.arguments += event.delta ?? ''
      continue
    }
    if (event.type === 'response.function_call_arguments.done' && event.call_id) {
      const call = pending.get(event.call_id)
      if (!call) continue
      pending.delete(event.call_id)
      request.onToolCall({
        id: event.call_id,
        name: event.name ?? call.name,
        input: toolInput(event.arguments ?? call.arguments),
      })
      continue
    }
    if (event.type === 'response.output_item.done' && event.item?.type === 'function_call') {
      const callId = event.item.call_id
      if (!callId) continue
      const call = pending.get(callId) ?? { name: event.item.name ?? '', arguments: '' }
      pending.delete(callId)
      request.onToolCall({
        id: callId,
        name: event.item.name ?? call.name,
        input: toolInput(event.item.arguments ?? call.arguments),
      })
      continue
    }
    if (event.type === 'error' || event.type === 'response.failed') {
      throw new Error(
        event.error?.message ?? event.response?.error?.message ?? 'Codex stream failed',
      )
    }
  }
}
