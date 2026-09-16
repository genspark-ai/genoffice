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
  listCodexModels,
  isAiOverloadedError,
  isAiQuotaExhaustedError,
  maxOutputTokensOf,
  streamForProvider,
} from '@genoffice/ai-provider'
import { fetchRemoteImage } from '@genoffice/electron-utils/remote-image'

import { parseDuckDuckGo, parseDuckDuckGoImages } from '@genoffice/agent-skills'
import {
  assessFileCoverage,
  buildDictionary,
  buildTranslationPrompt,
  fillDictionaryGaps,
  buildTranslateSystemPrompt,
  defaultOutputPath,
  extractTranslationText,
  isSupportedExtension,
  KnowledgeBase,
  resolveTranslateSkills,
  sharedMemory,
  PersistentTranslationMemory,
  SUPPORTED_EXTENSIONS,
  translateBatch,
  translateFile,
  translateOne,
  type KBEntry,
  type KBListFilters,
  type TerminologyPair,
} from '@genoffice/translation-core'

// Re-export the shared prompt / text helpers so existing tests and external
// callers don't need to know that the canonical home moved to
// `@genoffice/translation-core`.
export { buildTranslationPrompt, buildTranslateSystemPrompt, extractTranslationText }

// Shared translation knowledge base — mirrors LumosAI's translate-config
// 5-schema KB at ~/.genoffice/translation-kb.json. The load promise is
// memoised so every caller can `await ensureKbLoaded()`; that ordering matters
// because a fire-and-forget load would race a subsequent upsert and clobber
// the in-memory store with the (empty) on-disk state.
export const sharedKnowledgeBase = new KnowledgeBase()
let kbLoadPromise: Promise<void> | null = null
function ensureKbLoaded(): Promise<void> {
  if (!kbLoadPromise) {
    kbLoadPromise = sharedKnowledgeBase.load().catch((err: unknown) => {
      console.warn('[translation-kb] load failed:', err)
    })
  }
  return kbLoadPromise
}

// Persistent translation memory — the in-memory `sharedMemory` loses everything
// on restart, so any LLM translation the agent makes is wasted work on the next
// run. `translationMemory` writes per-pair JSON under
// `~/.genoffice/translation-memory/` and rehydrates on boot, so sentence-level
// reuse kicks in without the user having to wire anything up.
export const translationMemory = new PersistentTranslationMemory()
let tmLoadPromise: Promise<void> | null = null
function ensureMemoryLoaded(): Promise<void> {
  if (!tmLoadPromise) {
    tmLoadPromise = translationMemory.load().catch((err: unknown) => {
      console.warn('[translation-memory] load failed:', err)
    })
  }
  return tmLoadPromise
}

// Debounce flushes so a burst of translations doesn't hit the disk on every
// call; the underlying `isDirty` already prevents redundant work.
let flushTimer: NodeJS.Timeout | null = null
function scheduleMemoryFlush(delayMs = 250): void {
  if (flushTimer) return
  flushTimer = setTimeout(() => {
    flushTimer = null
    void translationMemory.flush().catch((err: unknown) => {
      console.warn('[translation-memory] flush failed:', err)
    })
  }, delayMs)
}

// Most recently generated `--dictionary`.
//
// The Settings pane builds a dictionary for a file, then the user usually wants
// to keep translating short snippets with the terminology they just curated.
// Rather than make them re-paste it, the snippet path reuses the last built
// dictionary (or an explicit `dictionaryPath`). Parsed lazily and cached by
// path so repeated snippet calls are free.
let lastDictionary: { path: string; pairs: TerminologyPair[] } | null = null

interface LoadedDictionary {
  path: string
  /** Terminology pairs, longest source first so multi-word terms win. */
  pairs: TerminologyPair[]
}

/**
 * Read a generated `{ "source": "target" }` dictionary off disk. Returns null
 * when the file is missing or malformed — the caller then translates without
 * the dictionary instead of failing the whole request.
 */
function loadDictionary(path: string): LoadedDictionary | null {
  if (lastDictionary?.path === path) return lastDictionary
  try {
    const parsed: unknown = JSON.parse(readFileSync(path, 'utf8'))
    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) return null
    const pairs: TerminologyPair[] = []
    for (const [source, target] of Object.entries(parsed as Record<string, unknown>)) {
      if (typeof target !== 'string') continue
      if (!source.trim() || !target.trim()) continue
      pairs.push({ source, target })
    }
    if (pairs.length === 0) return null
    pairs.sort((a, b) => b.source.length - a.source.length)
    lastDictionary = { path, pairs }
    return lastDictionary
  } catch {
    return null
  }
}

/** Remember a freshly written dictionary without a redundant disk read. */
function rememberDictionary(path: string, pairs: readonly TerminologyPair[]): void {
  if (pairs.length === 0) return
  lastDictionary = { path, pairs: [...pairs].sort((a, b) => b.source.length - a.source.length) }
}

/**
 * Cache the dictionary a builder call just wrote. The builder returns its
 * segments, so we avoid re-reading the file we just wrote; when a builder omits
 * them (older callers) we fall back to reading it back from disk.
 */
function rememberBuiltDictionary(result: {
  ok: boolean
  dictionaryPath?: string | undefined
  segments?: Array<{ source: string; target?: string | undefined }> | undefined
}): void {
  if (!result.ok || !result.dictionaryPath) return
  const pairs = (result.segments ?? [])
    .filter((seg) => seg.target !== undefined && seg.target.trim().length > 0)
    .map((seg) => ({ source: seg.source, target: seg.target as string }))
  if (pairs.length === 0) {
    loadDictionary(result.dictionaryPath)
    return
  }
  rememberDictionary(result.dictionaryPath, pairs)
}

// ----- settings persistence --------------------------------------------------

const AI_SETTINGS_FILE = join(DATA_DIR, 'ai-settings.json')

/** Human-readable provider label for error messages ("MiniMax", "OpenAI", …).
 *  Falls back to the raw id if the provider isn't in the registry. */
function providerLabel(id: string): string {
  const known: Record<string, string> = {
    genspark: 'Genspark',
    codex: 'Codex CLI',
    anthropic: 'Anthropic',
    gemini: 'Gemini',
    deepseek: 'DeepSeek',
    openai: 'OpenAI',
    kimi: 'Kimi',
    glm: 'GLM',
    qwen: 'Qwen',
    doubao: 'Doubao',
    minimax: 'MiniMax',
    xai: 'Grok',
    mistral: 'Mistral',
    openrouter: 'OpenRouter',
    requesty: 'Requesty',
    'opencode-zen': 'OpenCode Zen',
    'opencode-go': 'OpenCode Go',
    custom: 'Custom provider',
  }
  return known[id] ?? id
}

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
        : err instanceof AiCreditsError || isAiQuotaExhaustedError(err)
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

  /**
   * home:ai-capabilities — report per-feature availability so agents and the
   * UI can show an honest picture. Unlike the CLI's `capabilities` command
   * (which only counts keyed providers), this handler includes the zero-
   * config DuckDuckGo fallback for web/image search because the handlers
   * themselves fall back to it without asking the user to set anything up.
   */
  registerHandle('home:ai-capabilities', () => {
    type CapabilityReport = {
      available: boolean
      via: string
      fallback?: string
      configured: boolean
      note?: string
    }
    const report: Record<string, CapabilityReport> = {}

    // 1) web text search — DuckDuckGo's html endpoint is always available
    //    (no key, no rate-limit interaction); a keyed Serper/Tavily provider
    //    upgrades to "configured" once the user supplies a key.
    const searchSettings = aiSettings.search
    const searchProvider = searchSettings?.provider ?? 'genspark'
    const searchKeyed = searchProvider !== 'genspark'
      && !!searchSettings?.providers?.[searchProvider as 'serper' | 'tavily']?.apiKey
    report.search = {
      available: true,
      via: searchKeyed ? searchProvider : 'duckduckgo',
      fallback: searchKeyed ? 'duckduckgo' : undefined,
      configured: searchKeyed || aiSettings.gskToolsEnabled !== false,
      note: searchKeyed
        ? undefined
        : 'DuckDuckGo HTML endpoint works without a key; install a Serper/Tavily key for richer results.',
    }

    // 2) image search — same pattern. DuckDuckGo's /?iax=images endpoint is
    //    wired through ai:image-search and never requires a key.
    report.image_search = {
      available: true,
      via: searchKeyed ? searchProvider : 'duckduckgo',
      fallback: searchKeyed ? 'duckduckgo' : undefined,
      configured: searchKeyed || aiSettings.gskToolsEnabled !== false,
      note: searchKeyed
        ? undefined
        : 'DuckDuckGo image endpoint works without a key.',
    }

    // 3) image generation — needs a media provider with a key (or Genspark
    //    credits). DDG does not generate images.
    const mediaSettings = aiSettings.media
    const imageProvider = mediaSettings?.imageProvider ?? 'genspark'
    const imageKeyed = imageProvider !== 'genspark'
      && !!mediaSettings?.providers?.[imageProvider]?.apiKey
    report.image_generation = {
      available: imageKeyed || aiSettings.gskToolsEnabled !== false,
      via: imageKeyed
        ? imageProvider
        : (aiSettings.gskToolsEnabled === false ? 'none' : 'genspark'),
      configured: imageKeyed || aiSettings.gskToolsEnabled !== false,
      note: imageKeyed
        ? undefined
        : 'Add a key for OpenAI/Gemini/Doubao/etc. to generate images.',
    }

    // 4) media analysis — same shape as image generation.
    const analysisProvider = mediaSettings?.analysisProvider ?? 'genspark'
    const analysisKeyed = analysisProvider !== 'genspark'
      && !!mediaSettings?.providers?.[analysisProvider]?.apiKey
    report.media_analysis = {
      available: analysisKeyed || aiSettings.gskToolsEnabled !== false,
      via: analysisKeyed
        ? analysisProvider
        : (aiSettings.gskToolsEnabled === false ? 'none' : 'genspark'),
      configured: analysisKeyed || aiSettings.gskToolsEnabled !== false,
      note: analysisKeyed
        ? undefined
        : 'Add a key for OpenAI/Gemini/Claude/etc. to analyze images and video.',
    }

    return {
      ok: true,
      capabilities: report,
      provider: aiSettings.provider,
      gskToolsEnabled: aiSettings.gskToolsEnabled !== false,
    }
  })

  registerHandle('ai:codex-models', async () => {
    // Real lookup against the Codex CLI bridge; falls back to an empty
    // catalog when the CLI isn't on PATH (which is normal for the
    // standalone web build — Codex is an opt-in provider).
    try {
      return await listCodexModels(aiSettings.providers.codex?.cliPath)
    } catch (err) {
      return {
        models: [],
        defaultModel: '',
        error: err instanceof Error ? err.message : String(err),
      }
    }
  })

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
      // Quota/credit exhaustion first: retrying cannot help, and a 429 whose
      // body also carries a quota notice must not be reported as a transient
      // capacity blip (that message tells the user to do the one thing that
      // will never work).
      if (!result.ok && isAiQuotaExhaustedError(result.error)) {
        return {
          ok: false,
          errorCode: 'credits',
          error: `${providerLabel(provider)} quota exhausted — top up or switch provider in Settings → AI.`,
        } satisfies AiChatResponse
      }
      if (!result.ok && isAiOverloadedError(result.error)) {
        return {
          ok: false,
          errorCode: 'overloaded',
          error: 'The AI service is busy right now. Please retry shortly.',
        } satisfies AiChatResponse
      }
      return result as AiChatResponse
    } catch (err) {
      if (isAiQuotaExhaustedError(err)) {
        return {
          ok: false,
          errorCode: 'credits',
          error: `${providerLabel(provider)} quota exhausted — top up or switch provider in Settings → AI.`,
        } satisfies AiChatResponse
      }
      if (isAiOverloadedError(err)) {
        return {
          ok: false,
          errorCode: 'overloaded',
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
    await ensureKbLoaded()
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
      {
        provider,
        config: config as AiProviderConfig,
        memory: translationMemory,
        knowledgeBase: sharedKnowledgeBase,
      },
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
    await ensureKbLoaded()
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
      {
        provider,
        config: config as AiProviderConfig,
        memory: translationMemory,
        knowledgeBase: sharedKnowledgeBase,
      },
    )
  })
  scheduleMemoryFlush()

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
    await ensureMemoryLoaded()
    const response = translationMemory.saveMany({
      scene: req.scene ?? 'office',
      sourceLang: req.sourceLang ?? 'auto',
      targetLang: req.targetLang ?? 'auto',
      units,
    })
    await translationMemory.flush()
    return response
  })

  // home:translate-snippet — quick text-paste translation for the Settings pane.
  // Wraps translateOne with a string input instead of an editor range, and
  // forwards the customer's translation memory so memory-hit responses are
  // surfaced to the UI for the "saved N ms · cache" badge.
  registerHandle('home:translate-snippet', async (_event: unknown, request: unknown) => {
    const req = (request ?? {}) as {
      text?: string
      sourceLang?: string
      targetLang?: string
      customerName?: string
      /** Reuse a specific generated dictionary; defaults to the last one built. */
      dictionaryPath?: string
      /** Set false to translate without any dictionary terminology. */
      useDictionary?: boolean
      settings?: AiSettings
    }
    const text = (req.text ?? '').trim()
    if (!text) return { ok: false, error: 'home:translate-snippet expected non-empty `text`' }
    if (!req.targetLang) {
      return { ok: false, error: 'home:translate-snippet expected non-empty `targetLang`' }
    }
    const incoming = req.settings || aiSettings
    const provider = incoming.provider
    const config = incoming.providers?.[provider]
    if (!config) return { ok: false, error: `AI provider "${provider}" not configured` }
    await ensureKbLoaded()

    // Reuse the dictionary the user just built (or one they point at) so the
    // snippet and file paths agree on terminology.
    const dictionary =
      req.useDictionary === false
        ? null
        : req.dictionaryPath
          ? loadDictionary(req.dictionaryPath)
          : lastDictionary
    const dictionaryPairs = dictionary?.pairs ?? []
    const dictionarySources = new Set(dictionaryPairs.map((pair) => pair.source))

    const started = Date.now()
    const result = await translateOne(
      {
        instruction: text,
        sourceLang: req.sourceLang,
        targetLang: req.targetLang,
        ...(req.customerName ? { glossaryCategory: req.customerName } : {}),
      },
      {
        provider,
        config: config as AiProviderConfig,
        memory: translationMemory,
        knowledgeBase: sharedKnowledgeBase,
        ...(dictionaryPairs.length > 0 ? { dictionary: dictionaryPairs } : {}),
      },
    )
    scheduleMemoryFlush()
    // `matchedTerms` merges KB and dictionary hits; split them so the pane can
    // label the two provenances separately.
    const allTerms = result.matchedTerms ?? []
    const dictionaryHits = allTerms.filter((term) => dictionarySources.has(term))
    const kbTerms = allTerms.filter((term) => !dictionarySources.has(term))
    return {
      ok: result.ok,
      translation: result.translated ?? '',
      status: result.status,
      matchedTerms: kbTerms,
      dictionaryHits,
      dictionary: dictionary
        ? { path: dictionary.path, terms: dictionaryPairs.length, hits: dictionaryHits.length }
        : null,
      sourceLang: req.sourceLang,
      targetLang: req.targetLang,
      elapsedMs: Date.now() - started,
      error: result.error,
    }
  })

  // ai:translate-build-dictionary — KB + LLM produce the `--dictionary` the
  // upstream file handlers consume, so the user never hand-writes one.
  registerHandle('ai:translate-build-dictionary', async (_event: unknown, request: unknown) => {
    const req = (request ?? {}) as {
      inputPath?: string
      sourceLang?: string
      targetLang?: string
      outputPath?: string
      maxSegments?: number
      minChars?: number
      customerName?: string
      glossaryCategory?: string
      useLlm?: boolean
      settings?: AiSettings
    }
    if (!req.inputPath) {
      return { ok: false, error: 'ai:translate-build-dictionary expected a non-empty `inputPath`' }
    }
    if (!req.targetLang) {
      return { ok: false, error: 'ai:translate-build-dictionary expected a non-empty `targetLang`' }
    }
    const incoming = req.settings || aiSettings
    const provider = incoming.provider
    const config = incoming.providers?.[provider]
    const wantsLlm = req.useLlm !== false
    if (wantsLlm && !config) {
      return { ok: false, error: `AI provider "${provider}" not configured` }
    }
    await ensureKbLoaded()
    const built = await buildDictionary(
      {
        inputPath: req.inputPath,
        sourceLang: req.sourceLang ?? 'auto',
        targetLang: req.targetLang,
        ...(req.outputPath !== undefined ? { outputPath: req.outputPath } : {}),
        ...(req.maxSegments !== undefined ? { maxSegments: req.maxSegments } : {}),
        ...(req.minChars !== undefined ? { minChars: req.minChars } : {}),
        ...(req.customerName !== undefined ? { customerName: req.customerName } : {}),
        ...(req.glossaryCategory !== undefined ? { glossaryCategory: req.glossaryCategory } : {}),
        ...(req.useLlm !== undefined ? { useLlm: req.useLlm } : {}),
        dataDir: DATA_DIR,
      },
      {
        translateBatch: async (input) => {
          if (!config) return { ok: false, error: `AI provider "${provider}" not configured` }
          return translateBatch(input, {
            provider,
            config: config as AiProviderConfig,
            memory: translationMemory,
            knowledgeBase: sharedKnowledgeBase,
          })
        },
      },
    )
    rememberBuiltDictionary(built)
    return built
  })

  // ai:translate-file-auto — build the dictionary, then translate in one call.
  // This is the flow the UI uses; the two halves stay separately callable so a
  // user can review or hand-edit a dictionary before spending the file pass.
  registerHandle('ai:translate-file-auto', async (_event: unknown, request: unknown) => {
    const req = (request ?? {}) as {
      inputPath?: string
      outputPath?: string
      sourceLang?: string
      targetLang?: string
      customerName?: string
      glossaryCategory?: string
      scale?: number
      /** Reuse an existing dictionary instead of building a new one. */
      dictionaryPath?: string
      settings?: AiSettings
    }
    if (!req.inputPath) {
      return { ok: false, error: 'ai:translate-file-auto expected a non-empty `inputPath`' }
    }
    if (!req.targetLang) {
      return { ok: false, error: 'ai:translate-file-auto expected a non-empty `targetLang`' }
    }
    if (!isSupportedExtension(req.inputPath)) {
      return {
        ok: false,
        error: `Unsupported file type; expected one of ${SUPPORTED_EXTENSIONS.join(', ')}`,
      }
    }
    // Re-run path: the caller already has a dictionary (typically the extended
    // one from ai:translate-fill-gaps) and does not want it rebuilt — that would
    // throw away the gap-filling work. It is a pure dictionary rewrite, so it
    // needs no provider and runs even before one is configured.
    if (req.dictionaryPath) {
      const scored = await assessFileCoverage({
        inputPath: req.inputPath,
        dictionaryPath: req.dictionaryPath,
      })
      if (!scored.ok) {
        return { ok: false, stage: 'dictionary', error: scored.error ?? 'dictionary unreadable' }
      }
      const fileResult = await translateFile({
        inputPath: req.inputPath,
        ...(req.outputPath !== undefined ? { outputPath: req.outputPath } : {}),
        dictionaryPath: req.dictionaryPath,
        ...(req.scale !== undefined ? { scale: req.scale } : {}),
      })
      scheduleMemoryFlush()
      return {
        ...fileResult,
        stage: fileResult.ok ? 'done' : 'translate',
        dictionaryPath: req.dictionaryPath,
        dictionaryReused: true,
        coverage: scored.coverage,
      }
    }

    const incoming = req.settings || aiSettings
    const provider = incoming.provider
    const config = incoming.providers?.[provider]
    if (!config) return { ok: false, error: `AI provider "${provider}" not configured` }
    await ensureKbLoaded()

    const dict = await buildDictionary(
      {
        inputPath: req.inputPath,
        sourceLang: req.sourceLang ?? 'auto',
        targetLang: req.targetLang,
        ...(req.customerName !== undefined ? { customerName: req.customerName } : {}),
        ...(req.glossaryCategory !== undefined ? { glossaryCategory: req.glossaryCategory } : {}),
        dataDir: DATA_DIR,
      },
      {
        translateBatch: async (input) =>
          translateBatch(input, {
            provider,
            config: config as AiProviderConfig,
            memory: translationMemory,
            knowledgeBase: sharedKnowledgeBase,
          }),
      },
    )
    if (!dict.ok || !dict.dictionaryPath) {
      return { ok: false, stage: 'dictionary', error: dict.error ?? 'dictionary build failed' }
    }
    rememberBuiltDictionary(dict)

    const fileResult = await translateFile({
      inputPath: req.inputPath,
      ...(req.outputPath !== undefined ? { outputPath: req.outputPath } : {}),
      dictionaryPath: dict.dictionaryPath,
      ...(req.scale !== undefined ? { scale: req.scale } : {}),
    })
    scheduleMemoryFlush()
    return {
      ...fileResult,
      stage: fileResult.ok ? 'done' : 'translate',
      dictionaryPath: dict.dictionaryPath,
      dictionary: {
        kbEntries: dict.kbEntries,
        llmEntries: dict.llmEntries,
        missed: dict.missed,
        totalSegments: dict.totalSegments,
        elapsedMs: dict.elapsedMs,
        segments: dict.segments,
      },
      // How much of the file the dictionary actually reached. The handlers
      // report nothing useful for a zh -> en pass (their miss heuristic only
      // looks at Latin/kana text), so the answer is recomputed from the mined
      // segments before the file pass has even run.
      coverage: dict.coverage,
    }
  })

  // ai:translate-fill-gaps — translate whatever the dictionary missed and write
  // an extended dictionary. This is the "fill in the values and re-run" step the
  // format handlers print, done by the model instead of by hand. Only the
  // uncovered segments are sent, so the cost is proportional to the gap.
  registerHandle('ai:translate-fill-gaps', async (_event: unknown, request: unknown) => {
    const req = (request ?? {}) as {
      inputPath?: string
      sourceLang?: string
      targetLang?: string
      dictionaryPath?: string
      outputPath?: string
      maxSegments?: number
      minChars?: number
      customerName?: string
      glossaryCategory?: string
      settings?: AiSettings
    }
    if (!req.inputPath) {
      return { ok: false, error: 'ai:translate-fill-gaps expected a non-empty `inputPath`' }
    }
    if (!req.targetLang) {
      return { ok: false, error: 'ai:translate-fill-gaps expected a non-empty `targetLang`' }
    }
    const dictionaryPath = req.dictionaryPath ?? lastDictionary?.path
    if (!dictionaryPath) {
      return {
        ok: false,
        error: 'ai:translate-fill-gaps found no dictionary to extend; build one first',
      }
    }
    const incoming = req.settings || aiSettings
    const provider = incoming.provider
    const config = incoming.providers?.[provider]
    if (!config) return { ok: false, error: `AI provider "${provider}" not configured` }
    await ensureKbLoaded()
    const result = await fillDictionaryGaps(
      {
        inputPath: req.inputPath,
        sourceLang: req.sourceLang ?? 'auto',
        targetLang: req.targetLang,
        dictionaryPath,
        ...(req.outputPath !== undefined ? { outputPath: req.outputPath } : {}),
        ...(req.maxSegments !== undefined ? { maxSegments: req.maxSegments } : {}),
        ...(req.minChars !== undefined ? { minChars: req.minChars } : {}),
        ...(req.customerName !== undefined ? { customerName: req.customerName } : {}),
        ...(req.glossaryCategory !== undefined ? { glossaryCategory: req.glossaryCategory } : {}),
        dataDir: DATA_DIR,
      },
      {
        translateBatch: async (input) =>
          translateBatch(input, {
            provider,
            config: config as AiProviderConfig,
            memory: translationMemory,
            knowledgeBase: sharedKnowledgeBase,
          }),
      },
    )
    scheduleMemoryFlush()
    // The extended dictionary is now the one worth reusing.
    if (result.ok && result.dictionaryPath && result.added && result.added > 0) {
      loadDictionary(result.dictionaryPath)
    }
    return result
  })

  // ----- translation knowledge base CRUD --------------------------------------
  // Mirrors LumosAI's translate-config CLI: list / add / remove / save.
  // The UI (Settings -> AI -> Translation Knowledge) and the agent tools both
  // hit these handlers; the JSON file at ~/.genoffice/translation-kb.json is
  // the single source of truth.
  registerHandle('ai:translation-kb-list', async (_event: unknown, filters: unknown) => {
    await ensureKbLoaded()
    const f = (filters ?? {}) as KBListFilters
    return { ok: true, entries: sharedKnowledgeBase.list(f) }
  })

  registerHandle('ai:translation-kb-upsert', async (_event: unknown, entry: unknown) => {
    await ensureKbLoaded()
    if (!entry || typeof entry !== 'object') {
      return { ok: false, error: 'ai:translation-kb-upsert expected a KB entry object' }
    }
    const candidate = entry as Partial<KBEntry> & { id?: string }
    if (!candidate.id || typeof candidate.id !== 'string') {
      return { ok: false, error: 'KB entry must include a string `id`' }
    }
    try {
      sharedKnowledgeBase.upsert(candidate as KBEntry)
      await sharedKnowledgeBase.save()
      return { ok: true, entry: candidate }
    } catch (err) {
      return { ok: false, error: err instanceof Error ? err.message : String(err) }
    }
  })

  registerHandle('ai:translation-kb-remove', async (_event: unknown, id: unknown) => {
    await ensureKbLoaded()
    const targetId = String(id ?? '').trim()
    if (!targetId) return { ok: false, error: 'ai:translation-kb-remove expected a non-empty id' }
    const removed = sharedKnowledgeBase.remove(targetId)
    if (removed) await sharedKnowledgeBase.save()
    return { ok: true, removed }
  })

  registerHandle('ai:translation-kb-resolve', async (_event: unknown, request: unknown) => {
    await ensureKbLoaded()
    const req = (request ?? {}) as {
      sourceLang?: string
      targetLang?: string
      category?: string
      customerName?: string
    }
    if (!req.targetLang) {
      return { ok: false, error: 'ai:translation-kb-resolve expected non-empty `targetLang`' }
    }
    const resolved = sharedKnowledgeBase.resolve({
      sourceLang: req.sourceLang ?? 'auto',
      targetLang: req.targetLang,
      ...(req.category !== undefined ? { category: req.category } : {}),
      ...(req.customerName !== undefined ? { customerName: req.customerName } : {}),
    })
    return { ok: true, ...resolved }
  })

  registerHandle('ai:translation-kb-stats', async () => {
    await ensureKbLoaded()
    const all = sharedKnowledgeBase.list()
    const bySchema: Record<string, number> = {}
    for (const entry of all) {
      const key = ('sourceTerm' in entry && 'targetTerm' in entry) ? 'term'
        : ('forbiddenText' in entry) ? 'forbidden'
        : ('policy' in entry) ? 'brand'
        : ('description' in entry && 'name' in entry) ? 'styleRule'
        : 'customerPreference'
      bySchema[key] = (bySchema[key] ?? 0) + 1
    }
    return { ok: true, total: all.length, bySchema, dirty: sharedKnowledgeBase.isDirty() }
  })

  // ----- whole-file translation (PDF / XLS(X) / PPTX / DOCX) ----------------
  // Delegates to the upstream LumosAI `translate.py`, which dispatches by
  // extension. The dictionary is built by the caller (typically from the KB +
  // an LLM extraction pass) and passed as a path; without one the script is a
  // no-op reformatter, which is why we surface that in the result rather than
  // pretending the file was translated.
  registerHandle('ai:translate-file', async (_event: unknown, request: unknown) => {
    const req = (request ?? {}) as {
      inputPath?: string
      outputPath?: string
      dictionaryPath?: string
      scale?: number
      timeoutMs?: number
    }
    if (!req.inputPath) {
      return { ok: false, error: 'ai:translate-file expected a non-empty `inputPath`' }
    }
    if (!isSupportedExtension(req.inputPath)) {
      return {
        ok: false,
        error: `Unsupported file type; expected one of ${SUPPORTED_EXTENSIONS.join(', ')}`,
      }
    }
    return translateFile({
      inputPath: req.inputPath,
      ...(req.outputPath !== undefined ? { outputPath: req.outputPath } : {}),
      ...(req.dictionaryPath !== undefined ? { dictionaryPath: req.dictionaryPath } : {}),
      ...(req.scale !== undefined ? { scale: req.scale } : {}),
      ...(req.timeoutMs !== undefined ? { timeoutMs: req.timeoutMs } : {}),
    })
  })

  // ai:translate-dictionary-status — which dictionary snippet translation will
  // reuse. Without this the pane would show a bare "reuse dictionary" checkbox
  // and the user could not tell which file it refers to.
  registerHandle('ai:translate-dictionary-status', () => ({
    ok: true,
    dictionary: lastDictionary
      ? { path: lastDictionary.path, terms: lastDictionary.pairs.length }
      : null,
  }))

  registerHandle('ai:translate-file-status', () => {
    const location = resolveTranslateSkills()
    return {
      ok: true,
      available: location.source !== 'missing',
      source: location.source,
      skillDir: location.skillDir,
      pythonPath: location.pythonPath,
      supportedExtensions: [...SUPPORTED_EXTENSIONS],
    }
  })

  registerHandle('ai:translate-file-output-path', (_event: unknown, inputPath: unknown) => {
    const p = String(inputPath ?? '')
    if (!p) return { ok: false, error: 'expected a non-empty inputPath' }
    return { ok: true, outputPath: defaultOutputPath(p) }
  })

  registerHandle('ai:web-search', async (_event: unknown, query: unknown) => {
    // Zero-config web search via DuckDuckGo HTML. Same parser the agent uses,
    // so the UI sees the same hits the agent does. If the network is blocked
    // (e.g. inside an air-gapped corp firewall) we surface the error so the
    // UI can show "search unavailable" instead of silently returning nothing.
    const q = String(query ?? '').trim()
    if (q.length < 2) {
      return { query: q, results: [], error: 'query must be at least 2 characters' }
    }
    const started = Date.now()
    try {
      const url = `https://html.duckduckgo.com/html/?q=${encodeURIComponent(q)}&kl=us-en`
      const response = await fetch(url, {
        headers: {
          'User-Agent':
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15',
          'Accept': 'text/html',
          'Accept-Language': 'en-US,en;q=0.9',
        },
      })
      if (!response.ok) {
        return {
          query: q,
          results: [],
          error: `DuckDuckGo HTTP ${response.status} ${response.statusText}`,
        }
      }
      const html = await response.text()
      const hits = parseDuckDuckGo(html, 8)
      return {
        query: q,
        results: hits,
        source: 'duckduckgo',
        ms: Date.now() - started,
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err)
      return { query: q, results: [], error: `DuckDuckGo unreachable: ${message}` }
    }
  })

  registerHandle('ai:image-search', async (_event: unknown, query: unknown) => {
    // Zero-config image search via DuckDuckGo's image endpoint. Same parser
    // the agent uses; works without any API key.
    const q = String(query ?? '').trim()
    if (q.length < 2) {
      return { query: q, results: [], error: 'query must be at least 2 characters' }
    }
    const started = Date.now()
    try {
      const url = `https://duckduckgo.com/?q=${encodeURIComponent(q)}&iax=images&ia=images`
      const response = await fetch(url, {
        headers: {
          'User-Agent':
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 Safari/605.1.15',
          'Accept': 'text/html',
        },
      })
      if (!response.ok) {
        return {
          query: q,
          results: [],
          error: `DuckDuckGo HTTP ${response.status} ${response.statusText}`,
        }
      }
      const html = await response.text()
      const hits = parseDuckDuckGoImages(html, 8)
      return {
        query: q,
        results: hits,
        source: 'duckduckgo',
        ms: Date.now() - started,
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err)
      return { query: q, results: [], error: `DuckDuckGo unreachable: ${message}` }
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
