import type { AgentMessage, AgentToolCall, AgentToolDef } from '@genoffice/agent-core'
import { isAiOverloadedError } from './overload-error'
import { withOutputCapFallback } from './output-cap'
import { streamAnthropic } from './protocols/anthropic'
import { streamGemini } from './protocols/gemini'
import { streamOpenAiCompatible } from './protocols/openai-compatible'
import { streamOpenAiResponses } from './protocols/openai-responses'
import { streamCodexAppServer } from './codex-app-server'
import { isKeylessProvider } from './agy-meta'
import { streamAgy } from './agy-cli'
import { firstJsonObject, readToolCallBody, resolveToolName } from './agy-tool-repair'
import {
  ensureLocalVlmEndpoint,
  localVlmRequestConfig,
  shouldFallbackFromAgyToLocalVlm,
} from './local-vlm-fallback'
import { LOCAL_VLM_IDLE_TIMEOUT_MS, LOCAL_VLM_SAMPLING_DEFAULTS } from './local-vlm-meta'
import type { StreamCallbacks } from './protocols/shared'
import { withSystemAddendum } from './system-addendum'
import { getProviderAdapter, type AiProtocol } from './registry'
import type { AiProviderConfig, AiProviderId } from './types'

export { streamAnthropic } from './protocols/anthropic'
export { streamGemini } from './protocols/gemini'
export { streamOpenAiCompatible } from './protocols/openai-compatible'
export { streamOpenAiResponses } from './protocols/openai-responses'
export { AiCreditsError, sseLines } from './protocols/shared'
export type { StreamCallbacks } from './protocols/shared'

const CORE_DOC_TOOLS = new Set([
  // Docs tools
  'get_document_context',
  'read_blocks',
  'insert_content',
  'replace_blocks',
  'delete_blocks',
  'apply_ops',
  'write_document',
  'create_document',
  // Sheets tools
  'get_workbook_context',
  'read_range',
  'aggregate_range',
  'propose_operations',
  'read_cells',
  'find_cells',
  'select_range',
])

function selectEffectiveToolsForLocalVlm(tools: AgentToolDef[]): {
  effectiveTools: AgentToolDef[]
  isDocEditor: boolean
} {
  // If document editing tools are present, we are in Docs/Sheets editor:
  // preserve the document editing tools so AI can read/write the file directly!
  const docTools = tools.filter((t) => CORE_DOC_TOOLS.has(t.name))
  if (docTools.length > 0) {
    return { effectiveTools: docTools, isDocEditor: true }
  }

  // If lightweight tools (<= 8), keep them
  if (tools.length <= 8 && JSON.stringify(tools).length <= 4000) {
    return { effectiveTools: tools, isDocEditor: false }
  }

  // Heavy composite toolset (e.g. HomeChat with 23 tools / 18,000 tokens):
  // strip to prevent context overflow while keeping normal chat responsive
  return { effectiveTools: [], isDocEditor: false }
}

function pruneSystemForLocalVlm(system: string, isDocEditor: boolean): string {
  if (!system) return ''
  const trimmed = system.trim()
  // When in Docs/Sheets editor, preserve document editing rules and recipes!
  if (isDocEditor) return trimmed
  if (trimmed.length <= 2500) return trimmed
  const lines = trimmed.split('\n')
  const filtered = lines.filter((line) => {
    const t = line.trim()
    if (t.startsWith('- `') || t.startsWith('Tool:') || t.includes('json_schema')) return false
    return true
  })
  const res = filtered.join('\n').trim()
  return res.length > 3000 ? res.slice(0, 3000) : res
}

/**
 * Ops that are SAFE to auto-assign { scope: 'document' } when target is missing.
 * Only styling/formatting ops with no destructive side-effects belong here.
 * DESTRUCTIVE ops (deleteBlocks, moveBlocks, findReplace) are intentionally excluded:
 * they must have an explicit target from the model — if missing, we reject them
 * rather than risk silently deleting/moving the entire document.
 */
const SAFE_AUTO_TARGET_OPS = new Set([
  'setFont',
  'setMatchedFont',
  'setParagraphFormat',
  'setHeadingLevel',
  'setList',
  'clearList',
  'setImageProperties',
  'insertToc',
  'setParagraphAttrs',
  'setOutlineLevel',
  'stepIndent',
  'stepHangingIndent',
])

/**
 * All ops the model may legitimately name directly (without wrapping in apply_ops).
 * Recognized as valid tool names during XML parsing so they are not leaked as text.
 * Destructive ops ARE listed here (the model IS allowed to call them) but they must
 * supply an explicit target — normalizeToolCall enforces this below.
 */
const KNOWN_OPS = new Set([
  ...SAFE_AUTO_TARGET_OPS,
  'deleteBlocks',
  'moveBlocks',
  'findReplace',
])

/** Ops that must have an explicit, non-document-scoped target to prevent data loss. */
const DESTRUCTIVE_OPS = new Set(['deleteBlocks', 'moveBlocks', 'findReplace'])

function normalizeToolCall(
  name: string,
  input: Record<string, unknown>,
  /** Names the caller is actually allowed to invoke (from availableTools). */
  allowedNames: ReadonlySet<string>,
): { name: string; input: Record<string, unknown> } | null {
  // ── Blocker 4: enforce tool allowlist ──────────────────────────────────────
  // Only pass through tool names that the current session was granted.
  // KNOWN_OPS map to 'apply_ops', so if 'apply_ops' is not in allowedNames we
  // reject everything that would arrive there too.
  const resolvedName = KNOWN_OPS.has(name) ? 'apply_ops' : name
  if (!allowedNames.has(resolvedName)) return null

  // ── Fix read_blocks defaults ───────────────────────────────────────────────
  if (name === 'read_blocks') {
    const fixedInput = { ...input }
    if (fixedInput.startBlockIndex === undefined) fixedInput.startBlockIndex = 0
    if (fixedInput.endBlockIndex === undefined) fixedInput.endBlockIndex = 0
    return { name, input: fixedInput }
  }

  // ── Wrap bare op name into apply_ops ──────────────────────────────────────
  if (KNOWN_OPS.has(name)) {
    // ── Blocker 2: refuse destructive ops without an explicit target ───────
    if (DESTRUCTIVE_OPS.has(name)) {
      if (!input.target) {
        // Return null → parseToolCallBlock will treat it as unrecognized and
        // emit a text delta instead of silently destroying data.
        return null
      }
      return {
        name: 'apply_ops',
        input: { ops: [{ op: name, ...input }] },
      }
    }

    // Safe op: auto-assign document scope when target is missing.
    const target = input.target ?? { scope: 'document' }
    return {
      name: 'apply_ops',
      input: { ops: [{ op: name, target, ...input }] },
    }
  }

  // ── Normalize apply_ops with flattened / missing ops array ────────────────
  if (name === 'apply_ops') {
    if (!Array.isArray(input.ops)) {
      if (typeof input.op === 'string') {
        const opName = input.op as string
        // Reject destructive ops without target even here
        if (DESTRUCTIVE_OPS.has(opName) && !input.target) return null
        const target =
          SAFE_AUTO_TARGET_OPS.has(opName) && !input.target
            ? { scope: 'document' }
            : input.target
        return {
          name: 'apply_ops',
          input: { ops: [{ target, ...input }] },
        }
      }
      // Malformed apply_ops with no ops array and no op field — reject
      return null
    }

    // Normalize each op inside the array
    const fixedOps = (input.ops as Array<Record<string, unknown>>).map((op) => {
      if (!op || typeof op !== 'object' || typeof op.op !== 'string') return op
      const opName = op.op as string
      // Destructive ops must already have a target — do not auto-assign
      if (DESTRUCTIVE_OPS.has(opName)) return op
      // Safe ops: auto-assign document scope when missing
      if (SAFE_AUTO_TARGET_OPS.has(opName) && !op.target) {
        return { target: { scope: 'document' }, ...op }
      }
      return op
    })
    return { name: 'apply_ops', input: { ...input, ops: fixedOps } }
  }

  return { name, input }
}

function parseToolCallBlock(
  fullBlock: string,
  known: ReadonlySet<string>,
  allowedNames: ReadonlySet<string>,
): AgentToolCall | null {
  const inner = fullBlock
    .replace(/^<tool_call>/i, '')
    .replace(/<\/tool_call>$/i, '')
    .trim()

  // 1. XML function syntax: <function=name>...</function>
  const fnMatch = /<function=([a-zA-Z0-9_\-]+)>([\s\S]*?)<\/function>/i.exec(inner)
  if (fnMatch) {
    const rawName = fnMatch[1]!
    const body = fnMatch[2]?.trim() ?? ''
    // Only resolve to a known alias; never fall back to rawName for unknown tools.
    const resolved = resolveToolName(rawName, known)
    if (!resolved) return null
    let input: Record<string, unknown> = {}

    if (body) {
      const argMatch = /<arguments>([\s\S]*?)<\/arguments>/i.exec(body)
      const rawArgs = (argMatch ? argMatch[1]!.trim() : body)
      if (rawArgs) {
        try {
          const parsed = JSON.parse(rawArgs)
          if (typeof parsed === 'object' && parsed !== null && !Array.isArray(parsed)) {
            input = parsed as Record<string, unknown>
          }
        } catch {
          const jsonSub = firstJsonObject(rawArgs)
          if (jsonSub) {
            try {
              input = JSON.parse(jsonSub.json)
            } catch {}
          }
        }
      }
    }

    const normalized = normalizeToolCall(resolved, input, allowedNames)
    if (!normalized) return null

    return {
      id: `local-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`,
      name: normalized.name,
      input: normalized.input,
    }
  }

  // 2. JSON syntax: {"name": "...", "arguments": {...}}
  const obj = firstJsonObject(inner)
  if (obj) {
    const call = readToolCallBody(obj.json, known)
    if (call) {
      const normalized = normalizeToolCall(call.name, call.input, allowedNames)
      if (!normalized) return null
      return {
        id: `local-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`,
        name: normalized.name,
        input: normalized.input,
      }
    }
  }

  return null
}

function wrapWithToolCallFilter(
  cb: StreamCallbacks,
  availableTools: readonly AgentToolDef[],
): StreamCallbacks & { flush: () => void } {
  // allowedNames = exact tool names this session is permitted to invoke.
  // KNOWN_OPS map to 'apply_ops', so they are only permitted if 'apply_ops' is in the list.
  const allowedNames = new Set(availableTools.map((t) => t.name))

  // known = all name strings we scan for inside <tool_call> blocks (superset of allowedNames).
  // This lets us recognise KNOWN_OPS by name even though they will be remapped to apply_ops.
  const known = new Set(allowedNames)
  for (const name of [
    'read_blocks',
    'insert_content',
    'replace_blocks',
    'delete_blocks',
    'apply_ops',
    'get_document_context',
    'write_document',
    ...KNOWN_OPS,
  ]) {
    known.add(name)
  }

  let buffer = ''
  let inToolCall = false
  let toolCallEmitted = false

  const flushSafeBuffer = () => {
    if (buffer && !inToolCall) {
      const match = buffer.match(/<[t]?[o]?[o]?[l]?[_]?[c]?[a]?[l]?[l]?[>]?$/i)
      if (match) {
        const safe = buffer.slice(0, match.index)
        if (safe) cb.onDelta(safe)
        buffer = buffer.slice(match.index)
      } else {
        cb.onDelta(buffer)
        buffer = ''
      }
    }
  }

  const flush = () => {
    if (buffer) {
      cb.onDelta(buffer)
      buffer = ''
    }
    if (toolCallEmitted) {
      cb.onStopReason?.('tool_calls')
    }
  }

  return {
    ...cb,
    flush,
    onDelta: (chunk: string) => {
      buffer += chunk

      while (buffer.length > 0) {
        if (!inToolCall) {
          const openMatch = /<tool_call>/i.exec(buffer)
          if (!openMatch) {
            flushSafeBuffer()
            return
          }
          if (openMatch.index > 0) {
            cb.onDelta(buffer.slice(0, openMatch.index))
            buffer = buffer.slice(openMatch.index)
          }
          inToolCall = true
        }

        const closeMatch = /<\/tool_call>/i.exec(buffer)
        if (!closeMatch) {
          // Waiting for more chunks to close </tool_call>
          return
        }

        const fullBlockEnd = closeMatch.index + closeMatch[0].length
        const fullBlock = buffer.slice(0, fullBlockEnd)
        buffer = buffer.slice(fullBlockEnd)
        inToolCall = false

        const call = parseToolCallBlock(fullBlock, known, allowedNames)
        if (call) {
          toolCallEmitted = true
          cb.onToolCall?.(call)
        } else {
          cb.onDelta(fullBlock)
        }
      }
    },
    onStopReason: (reason) => {
      if (buffer) {
        cb.onDelta(buffer)
        buffer = ''
      }
      cb.onStopReason?.(toolCallEmitted ? 'tool_calls' : reason)
    },
  }
}

function trimMessagesForLocalVlm(messages: AgentMessage[]): AgentMessage[] {
  // Keep at most 6 recent messages to prevent context overflow on 8k context window
  const recent = messages.length > 6 ? messages.slice(-6) : messages
  return recent.map((m, idx) => {
    // Keep the latest user message intact
    if (idx === recent.length - 1 && m.role === 'user') return m
    if (m.role === 'user' || m.role === 'assistant') {
      if (m.text && m.text.length > 2000) {
        return {
          ...m,
          text: m.text.slice(0, 2000) + '... [truncated]',
        }
      }
    } else if (m.role === 'tool') {
      const trimmedResults = m.results.map((r) => {
        if (typeof r.output === 'string' && r.output.length > 1500) {
          return {
            ...r,
            output: r.output.slice(0, 1500) + '... [truncated]',
          }
        }
        return r
      })
      return { ...m, results: trimmedResults }
    }
    return m
  })
}

async function streamLocalVlmOpenAi(
  local: { baseUrl: string; model: string },
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
): Promise<void> {
  const localConfig = localVlmRequestConfig({ model: local.model })
  const options = {
    bodyExtras: { ...LOCAL_VLM_SAMPLING_DEFAULTS },
    idleTimeoutMs: LOCAL_VLM_IDLE_TIMEOUT_MS,
  }
  // CPU Optimization:
  // In Docs/Sheets editor: pass the document editing tools (read_blocks, insert_content,
  // replace_blocks...) so AI can read/write the file directly!
  // In HomeChat: strip heavy non-doc tools to avoid 18k token blowup while keeping prompt < 300 tokens.
  const { effectiveTools, isDocEditor } = selectEffectiveToolsForLocalVlm(tools)
  const slimSystem = pruneSystemForLocalVlm(system, isDocEditor)
  const windowedMessages = trimMessagesForLocalVlm(messages)
  const filteredCb = wrapWithToolCallFilter(cb, effectiveTools)

  try {
    await streamOpenAiCompatible(
      local.baseUrl,
      localConfig,
      slimSystem,
      windowedMessages,
      effectiveTools,
      maxTokens,
      filteredCb,
      options,
    )
    filteredCb.flush()
  } catch (err) {
    if (
      /exceed_context_size_error|exceeds the available context size|context.*exceed/i.test(
        err instanceof Error ? err.message : String(err),
      )
    ) {
      return await streamOpenAiCompatible(
        local.baseUrl,
        localConfig,
        '',
        messages.slice(-2),
        [],
        maxTokens,
        filteredCb,
        options,
      )
    }
    throw err
  }
}

/** route a streaming, tool-calling-capable turn by provider id */
async function streamForProviderInner(
  provider: AiProviderId,
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
): Promise<void> {
  if (provider === 'genspark') throw new Error('Genspark sign-in is disabled in this build')
  if (!isKeylessProvider(provider) && !config.model?.trim()) {
    throw new Error('Choose an AI model in Settings → AI Model')
  }
  if (provider === 'custom') {
    if (!config.baseUrl?.trim()) throw new Error('Set the Base URL in Settings → AI Model')
  } else if (!isKeylessProvider(provider) && !config.apiKey?.trim()) {
    throw new Error('Set an API key in Settings → AI Model')
  }
  const endpoint = getProviderAdapter(provider).resolveEndpoint(config)
  const { baseUrl } = endpoint
  if (endpoint.model) config = { ...config, model: endpoint.model }
  if (endpoint.protocol === 'codex-app-server') {
    return streamCodexAppServer(config, system, messages, tools, maxTokens, cb)
  }
  if (endpoint.protocol === 'local-vlm') {
    const local = await ensureLocalVlmEndpoint(cb.signal)
    return streamLocalVlmOpenAi(local, system, messages, tools, maxTokens, cb)
  }
  if (endpoint.protocol === 'agy-cli') {
    if (await shouldFallbackFromAgyToLocalVlm(config.cliPath)) {
      const local = await ensureLocalVlmEndpoint(cb.signal)
      return streamLocalVlmOpenAi(local, system, messages, tools, maxTokens, cb)
    }
    return streamAgy(config, system, messages, tools, maxTokens, cb)
  }
  // A model the user declared a ceiling for outranks the caller's budget; the
  // endpoint's own 400 would otherwise cost a billed round trip.
  const cap = endpoint.maxTokens !== undefined ? Math.min(maxTokens, endpoint.maxTokens) : maxTokens
  const protocol: Exclude<AiProtocol, 'codex-app-server' | 'agy-cli' | 'local-vlm'> =
    endpoint.protocol
  return withOutputCapFallback(baseUrl, config.model, cap, (tokens) => {
    switch (protocol) {
      case 'anthropic':
        return streamAnthropic(config, system, messages, tools, tokens, cb, baseUrl, {
          headers: endpoint.headers,
          idleTimeoutMs: endpoint.idleTimeoutMs,
        })
      case 'gemini':
        return streamGemini(config, system, messages, tools, tokens, cb, baseUrl, {
          omitTemperature: endpoint.omitTemperature,
          headers: endpoint.headers,
          idleTimeoutMs: endpoint.idleTimeoutMs,
        })
      case 'openai-compatible':
        return streamOpenAiCompatible(baseUrl, config, system, messages, tools, tokens, cb, {
          omitTemperature: endpoint.omitTemperature,
          useMaxCompletionTokens: endpoint.useMaxCompletionTokens,
          bodyExtras: endpoint.bodyExtras,
          echoReasoning: endpoint.echoReasoning,
          headers: endpoint.headers,
          idleTimeoutMs: endpoint.idleTimeoutMs,
        })
      case 'openai-responses':
        return streamOpenAiResponses(baseUrl, config, system, messages, tools, tokens, cb, {
          bodyExtras: endpoint.bodyExtras,
          headers: endpoint.headers,
          idleTimeoutMs: endpoint.idleTimeoutMs,
        })
    }
  })
}

export interface AiErrorDiagnostic {
  timestamp: string
  provider: AiProviderId
  status: number | null
  category: 'rate-limit' | 'overloaded' | 'auth' | 'network' | 'other'
}
let errorLogger: ((record: AiErrorDiagnostic) => void) | null = null
export function setAiErrorLogger(logger: (record: AiErrorDiagnostic) => void): void {
  errorLogger = logger
}

/** Record metadata only. Provider bodies can contain credentials or document text. */
export async function streamForProvider(
  ...args: Parameters<typeof streamForProviderInner>
): Promise<void> {
  try {
    // appended, never prepended: some callers recognise their own prompt by its first words
    const [provider, config, system, ...rest] = args
    await streamForProviderInner(provider, config, withSystemAddendum(system), ...rest)
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error)
    if (!(error instanceof Error && error.name === 'AbortError')) {
      const match = /(?:HTTP|status|code)[ :="']+(\d{3})\b/i.exec(message)
      const status = match ? Number(match[1]) : null
      const category =
        status === 429 || /quota|rate.?limit|resource.exhausted/i.test(message)
          ? 'rate-limit'
          : status === 503 ||
              status === 529 ||
              isAiOverloadedError(message) ||
              /unavailable|busy/i.test(message)
            ? 'overloaded'
            : status === 401 || status === 403
              ? 'auth'
              : /fetch failed|network|timeout|ECONN|ENOTFOUND/i.test(message)
                ? 'network'
                : 'other'
      try {
        errorLogger?.({ timestamp: new Date().toISOString(), provider: args[0], status, category })
      } catch {
        /* Diagnostics must never mask the provider failure. */
      }
    }
    throw error
  }
}
