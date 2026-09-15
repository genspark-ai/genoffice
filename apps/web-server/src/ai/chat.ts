/**
 * Core AI handlers — settings, login, chat, stream.
 *
 * Mirrors the Electron main-process surface so the same renderer code
 * (AgentLoop + createElectronTransport) works against the standalone
 * web-server without code changes. The streaming path now calls
 * `streamForProvider` from `@genoffice/ai-provider` — the same unified
 * entrypoint the Electron docs/sheets/slides apps use — so the web build
 * talks to real providers (MiniMax via OpenAI-compatible, Anthropic,
 * Gemini, etc.) instead of returning canned text.
 *
 * Settings live on disk in DATA_DIR/ai-settings.json so they survive
 * restarts and operator-driven `WEB_STATIC_ROOT` deploys.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { AI_STREAMS, DATA_DIR, registerHandle } from '../common/index'
import {
  AiCreditsError,
  AiTimeoutError,
  type AiChatRequest,
  type AiChatResponse,
  type AiProviderConfig,
  type AiProviderId,
  type AiSettings,
  type AiStreamChunk,
  chatForProvider,
  defaultAiSettings,
  isAiNetworkError,
  isAiOverloadedError,
  maxOutputTokensOf,
  streamForProvider,
} from '@genoffice/ai-provider'
import { fetchRemoteImage } from '@genoffice/electron-utils/remote-image'
import {
  buildTranslationPrompt,
  buildTranslateSystemPrompt,
  extractTranslationText,
  sharedMemory,
  translateBatch,
  translateOne,
} from '@genoffice/translation-core'

// Re-export the shared prompt / text helpers so existing tests and external
// callers don't need to know that the canonical home moved to
// `@genoffice/translation-core`.
export { buildTranslationPrompt, buildTranslateSystemPrompt, extractTranslationText }

// ----- settings persistence --------------------------------------------------

const AI_SETTINGS_FILE = join(DATA_DIR, 'ai-settings.json')

function loadSettings(): AiSettings {
  try {
    if (existsSync(AI_SETTINGS_FILE)) {
      const raw = readFileSync(AI_SETTINGS_FILE, 'utf8')
      const parsed = JSON.parse(raw) as AiSettings
      // re-merge on top of defaults so newly added providers appear without a wipe
      const def = defaultAiSettings()
      return {
        ...def,
        ...parsed,
        providers: { ...def.providers, ...(parsed.providers || {}) },
      }
    }
  } catch (err) {
    console.warn('[ai] failed to load ai-settings.json, falling back to defaults:', err)
  }
  // Seed from environment so the obvious deploy (set MINIMAX_API_KEY + start)
  // Just Works without anyone clicking through the settings UI.
  const env: Partial<Record<AiProviderId, string>> = {}
  if (process.env.MINIMAX_API_KEY) env.minimax = process.env.MINIMAX_API_KEY
  if (process.env.OPENAI_API_KEY) env.openai = process.env.OPENAI_API_KEY
  if (process.env.ANTHROPIC_API_KEY) env.anthropic = process.env.ANTHROPIC_API_KEY
  if (process.env.GEMINI_API_KEY) env.gemini = process.env.GEMINI_API_KEY
  if (process.env.DEEPSEEK_API_KEY) env.deepseek = process.env.DEEPSEEK_API_KEY
  if (process.env.KIMI_API_KEY) env.kimi = process.env.KIMI_API_KEY
  if (process.env.QWEN_API_KEY) env.qwen = process.env.QWEN_API_KEY
  if (process.env.DOUBAO_API_KEY) env.doubao = process.env.DOUBAO_API_KEY
  if (process.env.XAI_API_KEY) env.xai = process.env.XAI_API_KEY
  if (process.env.MISTRAL_API_KEY) env.mistral = process.env.MISTRAL_API_KEY
  if (process.env.OPENROUTER_API_KEY) env.openrouter = process.env.OPENROUTER_API_KEY
  const def = defaultAiSettings(env)
  // also pre-select a provider that actually has a key, so the very first
  // /api/ai/stream call hits the real LLM out of the box
  for (const k of Object.keys(env) as AiProviderId[]) {
    if (def.providers[k]?.apiKey) {
      def.provider = k
      break
    }
  }
  return def
}

function saveSettings(settings: AiSettings): void {
  try {
    mkdirSync(DATA_DIR, { recursive: true })
    writeFileSync(AI_SETTINGS_FILE, JSON.stringify(settings, null, 2), 'utf8')
  } catch (err) {
    console.warn('[ai] failed to persist ai-settings.json:', err)
  }
}

export let aiSettings: AiSettings = loadSettings()

// ----- shared streaming core -------------------------------------------------

export interface StreamSession {
  /** AbortController wired to the in-flight streamForProvider call */
  abort: AbortController
  /** how many chunks have been emitted (debug/observability) */
  chunks: number
}

export const AI_STREAM_SESSIONS = new Map<string, StreamSession>()

/**
 * Build the StreamCallbacks shape streamForProvider expects and forward
 * each event as an AiStreamChunk to the sender (IPC `event.sender.send`
 * in the IPC path; an HTTP SSE writer in the /api/ai/stream path).
 */
export interface AiStreamSink {
  send(chunk: AiStreamChunk): void
  onAbort?(controller: AbortController): void
}

export async function runProviderStream(
  settings: AiSettings,
  system: string,
  messages: Parameters<typeof streamForProvider>[3],
  tools: Parameters<typeof streamForProvider>[4],
  maxTokens: number | undefined,
  sink: AiStreamSink,
): Promise<void> {
  const provider = settings.provider
  const config = settings.providers?.[provider]
  if (!config) {
    sink.send({
      requestId: '',
      type: 'error',
      error: `AI provider "${provider}" not configured`,
    })
    return
  }
  if (provider !== 'genspark' && provider !== 'codex' && !config.apiKey) {
    sink.send({
      requestId: '',
      type: 'error',
      error: `No API key configured for provider "${provider}". Open Settings → AI to add one.`,
    })
    return
  }
  if (provider !== 'codex' && !config.model) {
    sink.send({ requestId: '', type: 'error', error: `No model selected for "${provider}".` })
    return
  }

  const controller = new AbortController()
  sink.onAbort?.(controller)

  let stopReason: string | undefined
  try {
    await streamForProvider(
      provider,
      config as AiProviderConfig,
      system,
      messages,
      tools,
      maxTokens ?? maxOutputTokensOf(settings),
      {
        signal: controller.signal,
        onDelta: (text) => sink.send({ requestId: '', type: 'delta', text }),
        onReasoningDelta: (text) => sink.send({ requestId: '', type: 'reasoning', text }),
        onToolCall: (toolCall) => sink.send({ requestId: '', type: 'tool-call', toolCall }),
        onStopReason: (reason) => {
          stopReason = reason
        },
        onActivity: () => {
          // wire-level keepalive so the renderer watchdog can tell a live turn
          sink.send({ requestId: '', type: 'ping' })
        },
      },
    )
    sink.send({ requestId: '', type: 'done', ...(stopReason ? { stopReason } : {}) })
  } catch (err) {
    if (controller.signal.aborted) {
      sink.send({ requestId: '', type: 'done' })
      return
    }
    sink.send({
      requestId: '',
      type: 'error',
      error: err instanceof Error ? err.message : String(err),
      ...(err instanceof AiTimeoutError
        ? { errorCode: 'timeout' as const }
        : err instanceof AiCreditsError
          ? { errorCode: 'credits' as const }
          : isAiNetworkError(err)
            ? { errorCode: 'network' as const }
            : isAiOverloadedError(err)
              ? { errorCode: 'overloaded' as const }
              : {}),
    })
  }
}

// ----- IPC handlers ----------------------------------------------------------

export function registerAiCoreHandlers(): void {
  registerHandle('ai:get-settings', () => aiSettings)
  registerHandle('ai:set-settings', (_event: unknown, settings: unknown) => {
    const next = settings as AiSettings
    if (!next || typeof next !== 'object') {
      throw new Error('ai:set-settings expected an AiSettings object')
    }
    aiSettings = {
      ...aiSettings,
      ...next,
      providers: { ...aiSettings.providers, ...(next.providers || {}) },
    }
    saveSettings(aiSettings)
    return { ok: true }
  })

  registerHandle('ai:gsk-login', () => ({
    loggedIn: true,
    email: 'web-user@genoffice.ai',
    credits: 1000,
  }))

  registerHandle('ai:gsk-status', (_event: unknown, withEmail?: unknown) => {
    if (withEmail === false) return { loggedIn: false }
    return { loggedIn: true, email: 'web-user@genoffice.ai' }
  })

  registerHandle('ai:log-run-failure', () => ({ ok: true }))

  registerHandle('ai:codex-models', () => ({ models: [], defaultModel: '' }))

  /**
   * ai:chat — non-streaming one-shot call. Mirrors the Electron
   * docs-main.ts:2987 handler so the shell's quick-prompt path works
   * identically against the web build. The renderer ships `{settings, system,
   * user}`; we resolve the active provider, gate on API key + model, and
   * route through `chatForProvider` from `@genoffice/ai-provider` (same
   * call the Electron side uses — no duplicate logic).
   */
  registerHandle('ai:chat', async (_event: unknown, request: unknown) => {
    const req = request as AiChatRequest | undefined
    if (!req || typeof req.user !== 'string') {
      throw new Error('ai:chat expected { settings, system, user }')
    }
    const incoming = req.settings || aiSettings
    const provider = incoming.provider
    const config = incoming.providers?.[provider]
    if (!config) {
      return {
        ok: false,
        error: `AI provider "${provider}" not configured`,
      } satisfies AiChatResponse
    }
    if (provider !== 'codex' && !config.apiKey) {
      return {
        ok: false,
        error:
          provider === 'genspark'
            ? 'Genspark account is not signed in. Sign in to use Genspark credits.'
            : `No API key configured for provider "${provider}". Open Settings → AI to add one.`,
      } satisfies AiChatResponse
    }
    if (provider !== 'codex' && !config.model) {
      return { ok: false, error: `No model selected for "${provider}".` } satisfies AiChatResponse
    }
    try {
      const result = await chatForProvider(
        provider,
        config as AiProviderConfig,
        req.system || '',
        req.user,
      )
      if (!result.ok && isAiOverloadedError(result.error)) {
        return {
          ok: false,
          error: 'The AI service is busy right now. Please retry shortly.',
        } satisfies AiChatResponse
      }
      return result as AiChatResponse
    } catch (err) {
      if (isAiOverloadedError(err)) {
        return {
          ok: false,
          error: 'The AI service is busy right now. Please retry shortly.',
        } satisfies AiChatResponse
      }
      return {
        ok: false,
        error: err instanceof Error ? err.message : String(err),
      } satisfies AiChatResponse
    }
  })

  /**
   * ai:translate — one-shot translate via the active provider. Reuses
   * `chatForProvider` with a hardened system prompt that forbids restyling.
   *
   * Input:  { instruction, sourceLang, targetLang, preserveFormat, range? }
   * Output: { ok, translated?, planId?, error? }
   *
   * Unlike the per-app skills (`ai:doc-write-*`, `ai:sheets-*`, `ai:slides-*`)
   * translate is real on the web build — the active LLM does the work.
   */
  function castEditorRange(raw: unknown): import('@genoffice/translation-core').EditorRange | null {
    if (!raw || typeof raw !== 'object') return null
    const r = raw as { from?: number; to?: number; scope?: string }
    const scope =
      r.scope === 'selection' ||
      r.scope === 'document' ||
      r.scope === 'paragraph' ||
      r.scope === 'cell' ||
      r.scope === 'table'
        ? r.scope
        : undefined
    return { from: r.from, to: r.to, scope }
  }

  registerHandle('ai:translate', async (_event: unknown, request: unknown) => {
    const req = (request ?? {}) as {
      instruction?: string
      sourceLang?: string
      targetLang?: string
      preserveFormat?: boolean
      range?: { from?: number; to?: number; scope?: string } | null
      memoryEnabled?: boolean
      qualityCheck?: boolean
      glossaryCategory?: string
      settings?: AiSettings
    }
    const incoming = req.settings || aiSettings
    const provider = incoming.provider
    const config = incoming.providers?.[provider]
    if (!config) {
      return { ok: false, error: `AI provider \"${provider}\" not configured` }
    }
    return translateOne(
      {
        instruction: req.instruction ?? '',
        sourceLang: req.sourceLang,
        targetLang: req.targetLang ?? '',
        preserveFormat: req.preserveFormat,
        range: castEditorRange(req.range),
        memoryEnabled: req.memoryEnabled,
        qualityCheck: req.qualityCheck,
        glossaryCategory: req.glossaryCategory,
      },
      { provider, config: config as AiProviderConfig, memory: sharedMemory },
    )
  })

  registerHandle('ai:translate-batch', async (_event: unknown, request: unknown) => {
    const req = (request ?? {}) as {
      units?: Array<{
        unitId?: string
        kind?: string
        sourceText?: string
        order?: number
        path?: string
        metadata?: Record<string, unknown>
        range?: { from?: number; to?: number; scope?: string } | null
      }>
      sourceLang?: string
      targetLang?: string
      preserveFormat?: boolean
      scene?: string
      memoryEnabled?: boolean
      qualityCheck?: boolean
      glossaryCategory?: string
    }
    const units = (req.units ?? []).map((u) => ({
      unitId: u.unitId ?? '',
      kind: (u.kind ?? 'paragraph') as
        'paragraph' | 'heading' | 'list-item' | 'table-cell' | 'document',
      sourceText: u.sourceText ?? '',
      order: u.order ?? 0,
      path: u.path,
      metadata: u.metadata,
      range: castEditorRange(u.range),
    }))
    const incoming = req as { settings?: AiSettings }
    const settings = (incoming as { settings?: AiSettings }).settings || aiSettings
    const provider = settings.provider
    const config = settings.providers?.[provider]
    if (!config) {
      return { ok: false, error: `AI provider "${provider}" not configured`, units: [] }
    }
    return translateBatch(
      {
        units,
        sourceLang: req.sourceLang,
        targetLang: req.targetLang ?? '',
        preserveFormat: req.preserveFormat,
        scene: req.scene,
        memoryEnabled: req.memoryEnabled,
        qualityCheck: req.qualityCheck,
        glossaryCategory: req.glossaryCategory,
      },
      { provider, config: config as AiProviderConfig, memory: sharedMemory },
    )
  })

  registerHandle('ai:save-translation-memory', async (_event: unknown, request: unknown) => {
    const req = (request ?? {}) as {
      scene?: string
      sourceLang?: string
      targetLang?: string
      units?: Array<{ unitId?: string; sourceText?: string; translatedText?: string }>
    }
    const units = (req.units ?? [])
      .filter((u) => u.sourceText && u.translatedText)
      .map((u) => ({
        unitId: u.unitId ?? '',
        sourceText: u.sourceText ?? '',
        translatedText: u.translatedText ?? '',
      }))
    return sharedMemory.saveMany({
      scene: req.scene ?? 'office',
      sourceLang: req.sourceLang ?? 'auto',
      targetLang: req.targetLang ?? 'auto',
      units,
    })
  })

  registerHandle('ai:web-search', async (_event: unknown, query: unknown) => {
    // Free web search has no API key configured in the default web build.
    // Surface that explicitly so the UI can render "search unavailable" rather
    // than silently returning empty results.
    return {
      query: String(query ?? ''),
      results: [],
      error:
        'web search requires a search-provider API key (Tavily/Serper). Configure in Settings → Search.',
    }
  })

  registerHandle('ai:image-search', async (_event: unknown, query: unknown) => {
    return {
      query: String(query ?? ''),
      results: [],
      error: 'image search requires a media-provider key. Configure in Settings → Media.',
    }
  })

  registerHandle('ai:fetch-image', async (_event: unknown, url: unknown) => {
    if (typeof url !== 'string' || url.length > 4096) return null
    try {
      const response = await fetchRemoteImage(url)
      if (!response?.ok || !response.body) return null
      const declared = Number(response.headers.get('content-length') ?? 0)
      if (declared > 20 * 1024 * 1024) return null
      const reader = response.body.getReader()
      const chunks: Buffer[] = []
      let total = 0
      for (;;) {
        const part = await reader.read()
        if (part.done) break
        total += part.value.byteLength
        if (total > 20 * 1024 * 1024) {
          await reader.cancel()
          return null
        }
        chunks.push(Buffer.from(part.value))
      }
      const contentType = response.headers.get('content-type') ?? ''
      const mime = contentType.includes('png')
        ? 'image/png'
        : contentType.includes('gif')
          ? 'image/gif'
          : 'image/jpeg'
      return { base64: Buffer.concat(chunks).toString('base64'), mime }
    } catch {
      return null
    }
  })

  /**
   * ai:stream IPC entry — used by the docs/sheets/slides renderers through
   * their web-bridge transport. The request shape matches what Electron's
   * docs-main.ts handler accepts: requestId, settings, system, messages,
   * tools. We use the request's settings (renderer-side override) when
   * present, falling back to the server's persisted settings.
   */
  registerHandle('ai:stream', async (event: unknown, request: unknown) => {
    const req = request as {
      requestId?: string
      sessionId?: string
      settings?: AiSettings
      system?: string
      messages?: Parameters<typeof streamForProvider>[3]
      tools?: Parameters<typeof streamForProvider>[4]
      maxTokens?: number
    }
    const requestId = req.requestId || `s-${Date.now()}`
    const settings = req.settings || aiSettings
    const system = req.system || ''
    const messages = req.messages || []
    const tools = req.tools || []
    const sender = (event as { sender?: { send?: (ch: string, ...args: unknown[]) => void } })
      ?.sender

    if (!sender?.send) {
      throw new Error('ai:stream requires an IPC sender (web bridge should provide one)')
    }

    const session = { abort: new AbortController(), chunks: 0 }
    AI_STREAMS.set(requestId, session as unknown as { chunks: string[]; abort: AbortController })
    AI_STREAM_SESSIONS.set(requestId, session)

    try {
      await runProviderStream(settings, system, messages, tools, req.maxTokens, {
        onAbort: (c) => {
          session.abort = c
        },
        send: (chunk) => {
          session.chunks++
          sender!.send?.('ai:stream-chunk', { ...chunk, requestId })
        },
      })
    } finally {
      AI_STREAMS.delete(requestId)
      AI_STREAM_SESSIONS.delete(requestId)
    }
    return { id: requestId }
  })

  registerHandle('ai:stream-cancel', (_event: unknown, requestId: unknown) => {
    const session = AI_STREAM_SESSIONS.get(String(requestId))
    if (session) {
      session.abort.abort()
      AI_STREAM_SESSIONS.delete(String(requestId))
    }
    return { ok: true }
  })
}
