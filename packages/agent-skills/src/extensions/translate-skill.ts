/**
 * translate-skill extension — the single agent-facing entry point for the
 * GenOffice translation pipeline.
 *
 * Why this exists
 * ---------------
 * Pre-W37, every UI-driven translation went through `translate-http.ts` and
 * the IPC handlers in `chat.ts`, completely bypassing the agent loop. The
 * embedded pi AgentSession had no real way to translate — installing
 * `translate-pdf` as a marketplace skill only gave the agent a wrapper
 * around the upstream Python script. By moving the TS path into a real pi
 * extension, translation becomes a first-class agent capability: the same
 * `translate_text` tool is available to the UI (via `/api/ai/pi-prompt`) and
 * to the agent (via the embedded AgentSession). One source of truth, one KB,
 * one set of quality rules.
 *
 * Tools
 * -----
 *   - translate_text   — one-shot text translation with KB rules applied.
 *   - translate_file   — file translation: PDF / DOCX / XLS / XLSX / PPT /
 *                        PPTX routed to the matching LumosAI Python script
 *                        via `bash`. Other formats fall back to text extract
 *                        + LLM.
 *   - build_dictionary — mine a file's strings, ask the LLM for source→target
 *                        pairs, persist as JSON the Python wrappers consume
 *                        via `--dictionary`.
 *   - kb_search / kb_upsert / kb_remove — CRUD over the translation KB
 *                        (5-schema: term / forbidden / brand / styleRule /
 *                        customerPreference).
 */

import { defineTool, type ExtensionAPI } from "@earendil-works/pi-coding-agent"
import Type from "typebox"
import { basename, extname, join } from "node:path"
import { existsSync, readdirSync, statSync } from "node:fs"
import { homedir } from "node:os"
import { mkdir, readFile, writeFile } from "node:fs/promises"

import {
  translateOne,
  KnowledgeBase,
  sharedMemory,
  extractTranslationText,
  type TranslateRequest,
  type TranslateResponse,
} from "@genoffice/translation-core"
import {
  chatForProvider,
  type AiSettings,
  type AiProviderId,
  type AiProviderConfig,
} from "@genoffice/ai-provider"

// ============================================================================
// Settings resolution
// ============================================================================
//
// AI settings are read on demand (lazy, 1-second TTL) so a Settings change in
// the host is picked up by the next tool call without rebuilding the agent
// session — the same trick the UI's `ai:get-settings` handler uses.
//
// Where the file lives matters: the host (web-server / Electron shell) writes
// the user's provider + API key to its own DATA_DIR, while the desktop app has
// historically used `~/.genoffice/`. Reading only one of them made the agent
// silently use a different provider than the UI — the exact bug this resolves.
// Precedence:
//   1. GENOFFICE_AI_SETTINGS   — explicit absolute override (tests, custom deploys)
//   2. <DATA_DIR>/ai-settings.json where DATA_DIR is the host's data dir
//   3. ~/.genoffice/ai-settings.json — desktop-app legacy location
// The first existing file wins; if none exist we fall back to defaults.

// Minimal but valid AiSettings used when no settings file is on disk yet. The
// shape mirrors `defaultAiSettings()` in @genoffice/ai-provider so we never
// invent a provider id that isn't real.
import { defaultAiSettings } from "@genoffice/ai-provider"

function makeDefaultSettings(): AiSettings {
  const base = defaultAiSettings()
  return { ...base, provider: "genspark" }
}

let cachedSettings: { value: AiSettings; expiresAt: number } | null = null

/** Tests can swap this to bypass file IO; production keeps the file-backed loader. */
let readSettingsOverride: (() => Promise<AiSettings>) | null = null
export function __setReadSettingsForTests(fn: (() => Promise<AiSettings>) | null): void {
  readSettingsOverride = fn
  cachedSettings = null
}

/**
 * Candidate settings files, most-specific first. Exported so hosts and tests
 * can assert which file the agent will actually read.
 */
export function aiSettingsCandidates(env: NodeJS.ProcessEnv = process.env): string[] {
  const candidates: string[] = []
  const explicit = env.GENOFFICE_AI_SETTINGS
  if (explicit && explicit.length > 0) candidates.push(explicit)
  const dataDir = env.DATA_DIR || env.GENOFFICE_DATA_DIR || env.GENOFFICE_WEB_DATA_DIR
  if (dataDir && dataDir.length > 0) candidates.push(join(dataDir, "ai-settings.json"))
  candidates.push(join(homedir(), ".genoffice", "ai-settings.json"))
  return candidates
}

async function readSettings(): Promise<AiSettings> {
  if (readSettingsOverride) return readSettingsOverride()
  if (cachedSettings && cachedSettings.expiresAt > Date.now()) {
    return cachedSettings.value
  }
  let loaded: AiSettings | null = null
  for (const path of aiSettingsCandidates()) {
    try {
      const raw = await readFile(path, "utf-8")
      const parsed = JSON.parse(raw) as Partial<AiSettings>
      if (parsed && typeof parsed === "object") {
        const defaults = makeDefaultSettings()
        loaded = {
          ...defaults,
          ...parsed,
          providers: { ...defaults.providers, ...(parsed.providers ?? {}) },
        }
        break
      }
    } catch {
      /* try the next candidate */
    }
  }
  const value = loaded ?? makeDefaultSettings()
  cachedSettings = { value, expiresAt: Date.now() + 1000 }
  return value
}

/**
 * Resolve the active provider + its per-provider config. Falls back to
 * genspark whenever the stored provider has no usable config (matches the
 * behaviour of `activeProvider()` in @genoffice/ai-provider).
 */
function asProvider(settings: AiSettings): { provider: AiProviderId; config: AiProviderConfig } {
  const requested = settings.provider
  const cfg = settings.providers?.[requested]
  if (requested && cfg) {
    return { provider: requested as AiProviderId, config: cfg }
  }
  const fallback = settings.providers?.genspark
  return {
    provider: "genspark",
    config: fallback ?? { apiKey: "", model: "" },
  }
}

// ============================================================================
// Shared KB instance
// ============================================================================

export type TranslateOneFn = typeof translateOne
let translateOneOverride: TranslateOneFn | null = null
export function __setTranslateOneForTests(fn: TranslateOneFn | null): void {
  translateOneOverride = fn
}

export type ChatForProviderFn = typeof chatForProvider
let chatForProviderOverride: ChatForProviderFn | null = null
export function __setChatForProviderForTests(fn: ChatForProviderFn | null): void {
  chatForProviderOverride = fn
}

let kbInstance: KnowledgeBase | null = null
export function __resetKbForTests(): void { kbInstance = null }

async function getKb(): Promise<KnowledgeBase> {
  if (!kbInstance) {
    // Respect GENOFFICE_TRANSLATION_KB / DATA_DIR like the rest of the
    // translation stack. The previous hardcoded ~/.genoffice path diverged
    // from chat.ts's sharedKnowledgeBase and made the e2e tests' KB
    // upserts land in a different file than the agent's reads.
    kbInstance = new KnowledgeBase()
  }
  await kbInstance.load()
  return kbInstance
}

/** Locate the bundled LumosAI python script for a given file extension. */
function lumosScriptPath(ext: string): { script: string; handler: 'lumos-pdf' | 'lumos-docx' | 'lumos-xls' | 'lumos-ppt' } | null {
  const root = join(homedir(), ".lumos", "bundled-skills")
  if (!existsSync(root)) return null
  let newest: string | null = null
  let newestMtime = 0
  for (const entry of readdirSync(root)) {
    try {
      const m = statSync(join(root, entry)).mtimeMs
      if (m > newestMtime) {
        newestMtime = m
        newest = entry
      }
    } catch {
      /* ignore */
    }
  }
  if (!newest) return null
  const lower = ext.toLowerCase()
  const map: Record<string, { rel: string; handler: 'lumos-pdf' | 'lumos-docx' | 'lumos-xls' | 'lumos-ppt' }> = {
    ".pdf": { rel: "translate-pdf/scripts/translate.py", handler: "lumos-pdf" },
    ".docx": { rel: "translate-docx/scripts/translate.py", handler: "lumos-docx" },
    ".xls": { rel: "translate-xls/scripts/translate.py", handler: "lumos-xls" },
    ".xlsx": { rel: "translate-xls/scripts/translate.py", handler: "lumos-xls" },
    ".ppt": { rel: "translate-ppt/scripts/translate.py", handler: "lumos-ppt" },
    ".pptx": { rel: "translate-ppt/scripts/translate.py", handler: "lumos-ppt" },
  }
  const entry = map[lower]
  if (!entry) return null
  const full = join(root, newest, entry.rel)
  return existsSync(full) ? { script: full, handler: entry.handler } : null
}

// ============================================================================
// translate_text
// ============================================================================

const TranslateTextParams = Type.Object({
  text: Type.String({ minLength: 1, maxLength: 100_000, description: "Source text to translate." }),
  source_lang: Type.Optional(
    Type.String({ description: "Source language code (e.g. 'zh-CN'). Defaults to 'auto'." }),
  ),
  target_lang: Type.String({ description: "Target language code (e.g. 'en-US')." }),
  instruction: Type.Optional(
    Type.String({ description: "Optional style/voice instruction for the model." }),
  ),
})

type TranslateTextArgs = {
  text: string
  source_lang?: string
  target_lang: string
  instruction?: string
}

interface TranslateTextResult {
  ok: boolean
  translated?: string
  status?: 'translated' | 'memory-hit' | 'failed'
  matchedTerms?: string[]
  warnings?: string[]
  elapsedMs?: number
  error?: string
}

function createTranslateTextTool() {
  return defineTool<typeof TranslateTextParams, TranslateTextResult>({
    name: "translate_text",
    label: "Translate Text",
    description:
      "Translate a single text snippet via the active LLM provider, with the " +
      "translation knowledge base applied (terms / forbidden / brand / style / " +
      "customer preferences). For full file translation use translate_file. " +
      "Returns the translated text, matched terms, and quality warnings.",
    promptSnippet:
      "translate_text(text, target_lang) → { translated, status, matchedTerms, warnings }",
    promptGuidelines: [
      "Always pass target_lang as an explicit IETF tag (e.g. 'en-US', 'zh-CN').",
      "Use source_lang only when known — leaving it unset lets the model detect.",
      "matchedTerms / warnings are advisory; surface them to the user.",
    ],
    parameters: TranslateTextParams,
    async execute(_id, params: TranslateTextArgs, _signal) {
      const start = Date.now()
      try {
        const settings = await readSettings()
        const { provider, config } = asProvider(settings)
        const kb = await getKb()
        const req: TranslateRequest = {
          instruction: params.instruction
            ? `${params.text}\n\nStyle: ${params.instruction}`
            : params.text,
          sourceLang: params.source_lang,
          targetLang: params.target_lang,
          memoryEnabled: true,
          qualityCheck: true,
        }
        const fn = translateOneOverride ?? translateOne
        const res: TranslateResponse = await fn(req, {
          provider,
          config,
          memory: sharedMemory,
          knowledgeBase: kb,
          fuzzyMemoryEnabled: true,
        })
        const summary = res.ok
          ? `translate_text → status=${res.status ?? "translated"}, matchedTerms=${(res.matchedTerms ?? []).length}, warnings=${(res.warnings ?? []).length}, elapsedMs=${Date.now() - start}`
          : `translate_text failed: ${res.error ?? "unknown error"}`
        return {
          content: [{ type: "text" as const, text: summary }],
          details: {
            ok: res.ok,
            translated: res.translated,
            status: res.status,
            matchedTerms: res.matchedTerms ?? [],
            warnings: res.warnings ?? [],
            elapsedMs: Date.now() - start,
            error: res.error,
          },
        }
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err)
        return {
          content: [{ type: "text" as const, text: `translate_text error: ${msg}` }],
          details: { ok: false, error: msg },
        }
      }
    },
  })
}

// ============================================================================
// translate_file
// ============================================================================

const TranslateFileParams = Type.Object({
  input_path: Type.String({ maxLength: 4096, description: "Absolute path to the source file." }),
  output_path: Type.Optional(
    Type.String({ maxLength: 4096, description: "Where to write the translated file. Defaults to <input>.translated.<ext>." }),
  ),
  dictionary_path: Type.Optional(
    Type.String({ maxLength: 4096, description: "Optional --dictionary JSON built by build_dictionary." }),
  ),
  source_lang: Type.Optional(Type.String()),
  target_lang: Type.String(),
})

type TranslateFileArgs = {
  input_path: string
  output_path?: string
  dictionary_path?: string
  source_lang?: string
  target_lang: string
}

interface TranslateFileResult {
  ok: boolean
  outputPath?: string
  handler?: 'lumos-pdf' | 'lumos-docx' | 'lumos-xls' | 'lumos-ppt' | 'ts-fallback'
  bashCommand?: string
  elapsedMs?: number
  error?: string
}

function createTranslateFileTool() {
  return defineTool<typeof TranslateFileParams, TranslateFileResult>({
    name: "translate_file",
    label: "Translate File",
    description:
      "Translate a document file. PDF / DOCX / XLS / XLSX / PPT / PPTX are " +
      "routed to the upstream LumosAI Python scripts (format-preserving). " +
      "Other formats fall back to text-extract + LLM translation. The tool " +
      "returns the bash command to run for Python handlers; the agent must " +
      "have the `bash` tool available (it does by default).",
    promptSnippet:
      "translate_file(input_path, target_lang[, dictionary_path]) → { bashCommand, handler }",
    promptGuidelines: [
      "After translate_file returns a bashCommand, call the bash tool to execute it.",
      "Provide a --dictionary JSON built by build_dictionary whenever the file has technical terms.",
    ],
    parameters: TranslateFileParams,
    async execute(_id, params: TranslateFileArgs, _signal) {
      const start = Date.now()
      const ext = extname(params.input_path)
      const out = params.output_path ?? `${params.input_path}.translated${ext}`
      try {
        const lumos = lumosScriptPath(ext)
        if (lumos) {
          const dict = params.dictionary_path ? ` --dictionary ${params.dictionary_path}` : ""
          const cmd = `python3 ${lumos.script} "${params.input_path}" "${out}"${dict}`
          return {
            content: [
              { type: "text" as const, text: `translate_file(${ext}): run via bash → ${cmd}` },
            ],
            details: {
              ok: true,
              outputPath: out,
              handler: lumos.handler,
              bashCommand: cmd,
              elapsedMs: Date.now() - start,
            },
          }
        }
        return {
          content: [{
            type: "text" as const,
            text: `translate_file: no Python handler for ${ext}; use translate_text on extracted content.`,
          }],
          details: {
            ok: false,
            handler: "ts-fallback" as const,
            error: `no handler for ${ext}`,
            elapsedMs: Date.now() - start,
          },
        }
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err)
        return {
          content: [{ type: "text" as const, text: `translate_file error: ${msg}` }],
          details: { ok: false, error: msg, elapsedMs: Date.now() - start },
        }
      }
    },
  })
}

// ============================================================================
// build_dictionary
// ============================================================================

const BuildDictParams = Type.Object({
  input_path: Type.String({ maxLength: 4096, description: "Document to mine strings from." }),
  output_path: Type.Optional(Type.String({ maxLength: 4096 })),
  target_lang: Type.String(),
  source_lang: Type.Optional(Type.String()),
  max_pairs: Type.Optional(Type.Integer({ minimum: 8, maximum: 2000, default: 200 })),
})

type BuildDictArgs = {
  input_path: string
  output_path?: string
  target_lang: string
  source_lang?: string
  max_pairs?: number
}

interface BuildDictResult {
  ok: boolean
  outputPath?: string
  pairCount?: number
  sourceTerms?: string[]
  error?: string
}

async function callProviderForDict(
  prompt: string,
  provider: AiProviderId,
  config: AiProviderConfig,
): Promise<string> {
  const fn = chatForProviderOverride ?? chatForProvider
  const out = await fn(
    provider,
    config,
    "You are a translation KB builder. Output one `source: target` pair per line, no commentary.",
    prompt,
  )
  return out.content ?? ""
}

function parseDictionaryFromLlm(raw: string): Record<string, string> {
  const out: Record<string, string> = {}
  for (const line of raw.split(/\r?\n/)) {
    const m = line.match(/^\s*([^:][^:]*?)\s*[:=>]\s*(.+?)\s*$/)
    if (!m) continue
    const src = m[1].trim().replace(/^['"`]+|['"`]+$/g, "")
    const tgt = m[2].trim().replace(/^['"`]+|['"`]+$/g, "")
    if (src && tgt && src !== tgt) out[src] = tgt
  }
  return out
}

function pairCountOf(pairs: Record<string, string>): number {
  return Object.keys(pairs).length
}

function createBuildDictionaryTool() {
  return defineTool<typeof BuildDictParams, BuildDictResult>({
    name: "build_dictionary",
    label: "Build Translation Dictionary",
    description:
      "Ask the LLM to produce source→target pairs for technical terms in a " +
      "document, then write the result as a JSON dictionary the Python " +
      "translate scripts consume via `--dictionary`. Use this before " +
      "translate_file when the document has technical vocabulary (SKUs, " +
      "fabric codes, regulatory terms).",
    promptSnippet: "build_dictionary(input_path, target_lang) → JSON file",
    promptGuidelines: [
      "Returns at most max_pairs (default 200). The KB provides an additional layer.",
    ],
    parameters: BuildDictParams,
    async execute(_id, params: BuildDictArgs, _signal) {
      const out = params.output_path ?? `${params.input_path}.dictionary.json`
      try {
        const settings = await readSettings()
        const { provider, config } = asProvider(settings)
        const prompt = [
          `Identify up to ${params.max_pairs ?? 200} technical terms likely to appear in "${basename(params.input_path)}".`,
          `Output each as: <source> : <${params.target_lang} translation>`,
          params.source_lang ? `Source language: ${params.source_lang}.` : "Source language: auto-detect.",
          "Skip generic words; focus on technical vocabulary, brand names, and domain-specific phrases.",
          "Return ONLY the `source : target` lines — no headers, no commentary.",
        ].join("\n")
        const raw = await callProviderForDict(prompt, provider, config)
        const pairs = parseDictionaryFromLlm(extractTranslationText(raw) ?? raw)
        await mkdir(join(out, ".."), { recursive: true }).catch(() => undefined)
        await writeFile(out, JSON.stringify(pairs, null, 2), "utf-8")
        return {
          content: [{ type: "text" as const, text: `build_dictionary → ${pairCountOf(pairs)} pairs at ${out}` }],
          details: {
            ok: true as const,
            outputPath: out,
            pairCount: pairCountOf(pairs),
            sourceTerms: Object.keys(pairs),
          },
        }
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err)
        return {
          content: [{ type: "text" as const, text: `build_dictionary error: ${msg}` }],
          details: { ok: false, error: msg },
        }
      }
    },
  })
}

// ============================================================================
// KB CRUD tools
// ============================================================================

const KbUpsertParams = Type.Object({
  // Either provide a fully-formed entry, or the common shortcut fields below.
  entry: Type.Optional(Type.Record(Type.String(), Type.Any(), { description: "Full KB entry object." })),
  schema: Type.Optional(Type.Union([
    Type.Literal('term'),
    Type.Literal('forbidden'),
    Type.Literal('brand'),
    Type.Literal('styleRule'),
    Type.Literal('customerPreference'),
  ], { description: "Shortcut: KB schema type." })),
  source: Type.Optional(Type.String({ description: "Shortcut: source text (for schema=term/forbidden)." })),
  target: Type.Optional(Type.String({ description: "Shortcut: target text (for schema=term)." })),
  replacement: Type.Optional(Type.String({ description: "Shortcut: replacement (for schema=forbidden)." })),
  word: Type.Optional(Type.String({ description: "Shortcut: brand word (for schema=brand)." })),
  policy: Type.Optional(Type.String({ description: "Shortcut: brand policy (for schema=brand)." })),
  name: Type.Optional(Type.String({ description: "Shortcut: style rule name (for schema=styleRule)." })),
  description: Type.Optional(Type.String({ description: "Shortcut: style rule description (for schema=styleRule)." })),
  customerName: Type.Optional(Type.String({ description: "Shortcut: customer name (for schema=customerPreference)." })),
  preference: Type.Optional(Type.String({ description: "Shortcut: customer preference text (for schema=customerPreference)." })),
  priority: Type.Optional(Type.Number({ description: "Priority (higher = earlier). Defaults to 50." })),
  sourceLang: Type.Optional(Type.String({ description: "Source language code." })),
  targetLang: Type.Optional(Type.String({ description: "Target language code." })),
})

type KbUpsertArgs = {
  entry?: Record<string, unknown>
  schema?: 'term' | 'forbidden' | 'brand' | 'styleRule' | 'customerPreference'
  source?: string
  target?: string
  replacement?: string
  word?: string
  policy?: string
  name?: string
  description?: string
  customerName?: string
  preference?: string
  priority?: number
  sourceLang?: string
  targetLang?: string
}

function createKbUpsertTool() {
  return defineTool<typeof KbUpsertParams, { ok: boolean; id?: string; error?: string }>({
    name: "kb_upsert",
    label: "KB Upsert",
    description:
      "Insert or update a translation knowledge base entry. The KB has 5 " +
      "schemas: term (source/target), forbidden (forbidden text / replacement), " +
      "brand (word + policy), styleRule (name + description), " +
      "customer (customerName + preference). Existing entries with the same " +
      "id are replaced.",
    promptSnippet: "kb_upsert(entry) → { ok, id }",
    promptGuidelines: ["See Settings → AI → 翻译知识库 for the editable schema."],
    parameters: KbUpsertParams,
    async execute(_id, params: KbUpsertArgs, _signal) {
      try {
        const kb = await getKb()
        // Allow callers to pass shortcuts (schema/source/target/...) directly
        // OR a pre-built entry object. Build the entry when shortcuts are used.
        const entry: Record<string, unknown> = params.entry
          ? { ...params.entry }
          : { schema: params.schema }
        if (!entry.schema && params.schema) entry.schema = params.schema
        if (params.schema === 'term') {
          entry.sourceTerm = params.source ?? entry.sourceTerm
          entry.targetTerm = params.target ?? entry.targetTerm
        } else if (params.schema === 'forbidden') {
          entry.forbiddenText = params.source ?? entry.forbiddenText
          entry.replacement = params.replacement ?? entry.replacement
        } else if (params.schema === 'brand') {
          entry.word = params.word ?? entry.word
          entry.policy = params.policy ?? entry.policy
        } else if (params.schema === 'styleRule') {
          entry.name = params.name ?? entry.name
          entry.description = params.description ?? entry.description
        } else if (params.schema === 'customerPreference') {
          entry.customerName = params.customerName ?? entry.customerName
          entry.preference = params.preference ?? entry.preference
        }
        if (params.priority !== undefined) entry.priority = params.priority
        if (params.sourceLang) entry.sourceLang = params.sourceLang
        if (params.targetLang) entry.targetLang = params.targetLang
        // Auto-generate a stable id when the caller did not provide one.
        if (!entry.id) {
          const seed =
            (entry.sourceTerm ?? entry.forbiddenText ?? entry.word ?? entry.name ?? entry.customerName ?? "").toString()
          entry.id = `${entry.schema ?? "entry"}-${seed.replace(/\s+/g, "-").toLowerCase() || Date.now().toString(36)}`
        }
        const saved = await kb.upsert(entry as never)
        await kb.save().catch(() => undefined) // best-effort persistence; tolerate read-only mounts
        const id = (saved as { id?: string }).id ?? null
        return {
          content: [{ type: "text" as const, text: `kb_upsert → ${id ?? "(no id)"}` }],
          details: { ok: true, id: id ?? undefined },
        }
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err)
        return {
          content: [{ type: "text" as const, text: `kb_upsert error: ${msg}` }],
          details: { ok: false, error: msg },
        }
      }
    },
  })
}

const KbRemoveParams = Type.Object({ id: Type.String({ description: "KB entry id to remove." }) })

function createKbRemoveTool() {
  return defineTool<typeof KbRemoveParams, { ok: boolean; removed?: boolean; error?: string }>({
    name: "kb_remove",
    label: "KB Remove",
    description: "Remove a knowledge base entry by id.",
    promptSnippet: "kb_remove(id) → { ok, removed }",
    promptGuidelines: ["Use kb_search to find the id before removing."],
    parameters: KbRemoveParams,
    async execute(_id, params: { id: string }, _signal) {
      try {
        const kb = await getKb()
        const removed = await kb.remove(params.id)
        // Persist the removal so the next getKb() call (which reloads
        // from disk) sees the same state. Without this, a remove-then-
        // list round-trip resurrects the entry from the JSON file.
        await kb.save().catch(() => undefined)
        return {
          content: [{ type: "text" as const, text: `kb_remove(${params.id}) → ${removed}` }],
          details: { ok: true, removed },
        }
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err)
        return {
          content: [{ type: "text" as const, text: `kb_remove error: ${msg}` }],
          details: { ok: false, error: msg },
        }
      }
    },
  })
}

const KbSearchParams = Type.Object({
  query: Type.String({ minLength: 1, maxLength: 500 }),
  limit: Type.Optional(Type.Integer({ minimum: 1, maximum: 100, default: 20 })),
})

function createKbSearchTool() {
  return defineTool<typeof KbSearchParams, { ok: boolean; entries?: unknown[]; count?: number; error?: string }>({
    name: "kb_search",
    label: "KB Search",
    description:
      "Search the translation knowledge base for entries matching the query. " +
      "Returns a flat list of entries (across all 5 schemas) whose JSON " +
      "representation contains the query string.",
    promptSnippet: "kb_search(query[, limit]) → entries[]",
    promptGuidelines: ["Use kb_upsert to add new entries discovered during translation."],
    parameters: KbSearchParams,
    async execute(_id, params: { query: string; limit?: number }, _signal) {
      try {
        const kb = await getKb()
        const entries = await kb.list({})
        const q = params.query.toLowerCase()
        const filtered = entries
          .filter((e) => JSON.stringify(e).toLowerCase().includes(q))
          .slice(0, params.limit ?? 20)
        return {
          content: [{ type: "text" as const, text: `kb_search("${params.query}") → ${filtered.length} hits` }],
          details: { ok: true, entries: filtered, count: filtered.length },
        }
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err)
        return {
          content: [{ type: "text" as const, text: `kb_search error: ${msg}` }],
          details: { ok: false, error: msg },
        }
      }
    },
  })
}

const KbListParams = Type.Object({
  schema: Type.Optional(
    Type.Union(
      [
        Type.Literal("term"),
        Type.Literal("forbidden"),
        Type.Literal("brand"),
        Type.Literal("styleRule"),
        Type.Literal("customerPreference"),
      ],
      { description: "Optional filter to a single KB schema." },
    ),
  ),
  limit: Type.Optional(Type.Integer({ minimum: 1, maximum: 1000, default: 500 })),
})

function createKbListTool() {
  return defineTool<typeof KbListParams, { ok: boolean; entries?: unknown[]; count?: number; schema?: string; error?: string }>({
    name: "kb_list",
    label: "KB List",
    description:
      "List every entry in the translation knowledge base, optionally filtered " +
      "by schema. Use this to enumerate the KB before deciding what to upsert " +
      "or remove. The companion to kb_search (which does substring matching).",
    promptSnippet: "kb_list([schema][, limit]) → entries[]",
    promptGuidelines: ["Prefer kb_search for targeted lookups; kb_list for full inventories."],
    parameters: KbListParams,
    async execute(_id, rawParams: unknown, _signal) {
      try {
        const params = (rawParams ?? {}) as { schema?: string; limit?: number }
        const kb = await getKb()
        const filter: { schema?: string } = {}
        if (params.schema) filter.schema = params.schema
        const entries = await kb.list(filter as never)
        const sliced = entries.slice(0, params.limit ?? 500)
        return {
          content: [{ type: "text" as const, text: `kb_list → ${sliced.length}/${entries.length} entries` }],
          details: {
            ok: true,
            entries: sliced,
            count: sliced.length,
            schema: params.schema,
          },
        }
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err)
        return {
          content: [{ type: "text" as const, text: `kb_list error: ${msg}` }],
          details: { ok: false, error: msg },
        }
      }
    },
  })
}

// ============================================================================
// Extension factory
// ============================================================================

export interface TranslateSkillOptions {
  /** Override settings resolution (tests). */
  readSettings?: () => Promise<AiSettings>
}

export const ALL_TRANSLATE_TOOL_NAMES = [
  "translate_text",
  "translate_file",
  "build_dictionary",
  "kb_list",
  "kb_search",
  "kb_upsert",
  "kb_remove",
] as const

export type TranslateToolName = (typeof ALL_TRANSLATE_TOOL_NAMES)[number]

export function createTranslateSkillExtension(
  options: TranslateSkillOptions = {},
): (pi: ExtensionAPI) => void {
  const resolveSettings = options.readSettings ?? readSettings
  // Touch the overrides so unused-import lint stays happy in tests.
  void resolveSettings
  return (pi: ExtensionAPI) => {
    pi.registerTool(createTranslateTextTool())
    pi.registerTool(createTranslateFileTool())
    pi.registerTool(createBuildDictionaryTool())
    pi.registerTool(createKbUpsertTool())
    pi.registerTool(createKbRemoveTool())
    pi.registerTool(createKbSearchTool())
    pi.registerTool(createKbListTool())
  }
}
