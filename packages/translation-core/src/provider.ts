import { type AiProviderConfig, type AiProviderId } from '@genoffice/ai-provider'

import {
  buildTranslationPrompt,
  buildTranslateSystemPrompt,
  extractTranslationText,
  normalizeSourceLang,
} from './prompt'
import { assessBatchQuality, warningsFor } from './quality'
import { TranslationMemory } from './memory'
import type { MemoryEntry } from './memory'
import type { KnowledgeBase } from './knowledge-base'
import type {
  TranslateBatchRequest,
  TranslateBatchResponse,
  TranslateBatchUnitResult,
  TranslateRequest,
  TranslateResponse,
} from './types'
import { callLlm } from './llm-client'  // W9: seam between translation-core and the underlying LLM SDK
import {
  applyTerminology,
  matchTermsInSource,
  resolveKbForCall,
  terminologyPairs,
  type TerminologyPair,
} from './kb-rules'

/**
 * `translateOne` / `translateBatch` are the public call surface shared by the
 * Electron main-process handlers and the web-server. They wrap
 * `chatForProvider` with the hardened prompt and quality checks so every call
 * site gets the same behaviour and the same error mapping.
 *
 * The shared instance is a singleton so consecutive calls (the typical
 * retranslation flow) hit the in-memory TM.
 */
/** Subset of {@link TranslationMemory} the translator actually calls —
 *  the in-memory and persistent implementations both satisfy it. */
export interface TranslationMemoryLike {
  lookup(sourceLang: string, targetLang: string, sourceText: string): MemoryEntry | null
  save(entry: Omit<MemoryEntry, 'updatedAt'>): void
}

export const sharedMemory = new TranslationMemory()

export interface TranslateOneOptions {
  provider: AiProviderId
  config: AiProviderConfig
  /** Optional external memory; falls back to the shared in-memory TM.
   *  Accepts any object that exposes the methods the translator actually
   *  uses (`lookup` and `save`); the in-memory {@link TranslationMemory} and
   *  the file-backed {@link PersistentTranslationMemory} both satisfy this. */
  memory?: TranslationMemoryLike
  /**
   * Optional translation knowledge base. When provided, the resolved rules
   * (term / forbidden / brand / style / customer preferences) are appended
   * to the system prompt so the model respects the user's house style.
   * Hosts construct one via `new KnowledgeBase({ filePath })` and call
   * `.load()` on startup; the same instance is safe to share across
   * concurrent translate calls.
   */
  knowledgeBase?: KnowledgeBase
  /**
   * When the supplied `memory` exposes `fuzzyLookup`, fall back to a
   * similarity-based hit when no exact entry is found. Defaults to false so
   * the existing one-shot translation flow stays byte-for-byte identical.
   */
  fuzzyMemoryEnabled?: boolean
  /**
   * Extra mandatory `source -> target` pairs on top of the KB terms — the
   * generated `--dictionary` for the document being worked on. Hosts pass the
   * pairs (not the path) so translation-core stays filesystem-free; the
   * web-server resolves `ai:translate-build-dictionary` output into this slot.
   *
   * They are rendered into the system prompt, are matched against the source
   * text for `matchedTerms`, and are enforced on the model output exactly like
   * KB terms.
   */
  dictionary?: readonly TerminologyPair[] | undefined
}

/**
 * Translate a single chunk (selection / single paragraph). Returns the same
 * shape the docs `aiTranslate` IPC contract uses, so the caller can return the
 * value verbatim from the handler.
 */
export async function translateOne(
  request: TranslateRequest,
  opts: TranslateOneOptions,
): Promise<TranslateResponse> {
  const sourceText = (request.instruction ?? '').trim()
  const targetLang = (request.targetLang ?? '').trim()
  if (!sourceText) return { ok: false, error: 'ai:translate expected non-empty `instruction`' }
  if (!targetLang) return { ok: false, error: 'ai:translate expected non-empty `targetLang`' }
  if (!opts.config) {
    return { ok: false, error: `AI provider "${opts.provider}" not configured` }
  }
  if (opts.provider !== 'codex' && opts.provider !== 'genspark' && !opts.config.apiKey) {
    return {
      ok: false,
      error: `No API key configured for provider "${opts.provider}". Open Settings → AI to add one.`,
    }
  }
  if (opts.provider !== 'codex' && !opts.config.model) {
    return { ok: false, error: `No model selected for "${opts.provider}".` }
  }
  const memory = request.memoryEnabled === false
    ? null
    : (opts.memory ?? sharedMemory)
  const sourceLang = normalizeSourceLang(request.sourceLang)
  const preserveFormat = request.preserveFormat !== false

  // Resolve terminology once per call: KB mandatory terms plus whatever
  // dictionary the host layered on top. Both drive the prompt and the
  // `matchedTerms` the UI badges off, so they are computed before any early
  // return — a memory hit applied the same terms.
  const { pairs: termPairs, dictionaryTerms } = resolveTerminology({
    request,
    opts,
    sourceText,
    sourceLang,
    targetLang,
  })
  const matchedTerms = matchTermsInSource(sourceText, termPairs)
  const withTerms = matchedTerms.length > 0 ? { matchedTerms } : {}

  const hit = memory?.lookup(sourceLang, targetLang, sourceText)
  if (hit) {
    return {
      ok: true,
      translated: hit.translatedText,
      planId: `translate-memory-${Date.now().toString(36)}`,
      sourceLang,
      targetLang,
      preserveFormat,
      status: 'memory-hit',
      ...withTerms,
    }
  }
  // Optional fuzzy fallback — only when the host opts in by passing
  // `fuzzyMemoryEnabled: true` and the memory exposes `fuzzyLookup`.
  // PersistentTranslationMemory ships this method; the plain TranslationMemory
  // does not, so the lookup stays exact-match for legacy callers.
  if (opts.fuzzyMemoryEnabled && memory && typeof (memory as { fuzzyLookup?: unknown }).fuzzyLookup === 'function') {
    const fuzzyHit = (memory as unknown as { fuzzyLookup: (s: string, t: string, x: string) => { translatedText: string; confidence: number } | null }).fuzzyLookup(sourceLang, targetLang, sourceText)
    if (fuzzyHit) {
      return {
        ok: true,
        translated: fuzzyHit.translatedText,
        planId: `translate-fuzzy-${Date.now().toString(36)}`,
        sourceLang,
        targetLang,
        preserveFormat,
        status: 'memory-hit',
        warnings: [`fuzzy-match:${fuzzyHit.confidence.toFixed(2)}`],
        ...withTerms,
      }
    }
  }

  const system = buildTranslateSystemPrompt({
    sourceLang,
    targetLang,
    preserveFormat,
    glossaryCategory: request.glossaryCategory,
    ...(opts.knowledgeBase ? { knowledgeBase: opts.knowledgeBase } : {}),
    ...(dictionaryTerms.length > 0 ? { dictionaryTerms } : {}),
  })
  const metadata: Record<string, string> = {}
  if (request.glossaryCategory) metadata.glossaryCategory = request.glossaryCategory
  if (request.qualityCheck !== undefined) metadata.qualityCheck = String(request.qualityCheck)

  const result = await callLlm({
    provider: opts.provider,
    config: opts.config,
    systemPrompt: system,
    userPrompt: buildTranslationPrompt(sourceText),
    ...(Object.keys(metadata).length > 0 ? { metadata } : {}),
  })

  if (!result.ok) {
    return {
      ok: false,
      error: result.overloaded
        ? 'AI service is busy — please retry shortly.'
        : typeof result.error === 'string'
          ? result.error
          : 'Translation failed',
    }
  }
  const extracted = extractTranslationText(result.content ?? '')
  if (!extracted) {
    return { ok: false, error: 'Translation response did not contain final text.' }
  }
  // Enforce the mandatory terms the model may have left in the source language.
  const translated = applyTerminology(extracted, termPairs)
  if (memory) memory.save({ sourceLang, targetLang, sourceText, translatedText: translated })
  return {
    ok: true,
    translated,
    planId: `translate-${Date.now().toString(36)}`,
    sourceLang,
    targetLang,
    preserveFormat,
    status: 'translated',
    ...withTerms,
  }
}

/**
 * Build the per-call terminology for {@link translateOne}.
 *
 * `pairs` is everything that can rewrite the output (KB mandatory terms first,
 * then host dictionary pairs). `dictionaryTerms` is the subset the prompt
 * actually needs: the pairs whose source appears in this specific `sourceText`,
 * so a 400-segment dictionary does not blow up the system prompt for a one-line
 * snippet.
 */
function resolveTerminology(input: {
  request: TranslateRequest
  opts: TranslateOneOptions
  sourceText: string
  sourceLang: string
  targetLang: string
}): { pairs: TerminologyPair[]; dictionaryTerms: TerminologyPair[] } {
  const { request, opts, sourceText, sourceLang, targetLang } = input
  const fromKb = opts.knowledgeBase
    ? terminologyPairs(
        resolveKbForCall(opts.knowledgeBase, {
          sourceLang,
          targetLang,
          ...(request.glossaryCategory !== undefined ? { category: request.glossaryCategory } : {}),
        }),
      )
    : []
  const fromDictionary = opts.dictionary ?? []
  const dictionaryTerms = fromDictionary.filter((pair) => sourceText.includes(pair.source))
  return { pairs: [...fromKb, ...fromDictionary], dictionaryTerms }
}

/**
 * Terminology for a batch call. The dictionary subset injected into the prompt
 * is per-unit (see {@link resolveTerminology}); this only needs the pairs used
 * for `matchedTerms` and output enforcement.
 */
function terminologyForBatch(
  request: TranslateBatchRequest,
  opts: TranslateOneOptions,
  sourceLang: string,
  targetLang: string,
): TerminologyPair[] {
  return resolveTerminology({
    request: { instruction: '', targetLang },
    opts,
    sourceText: '',
    sourceLang,
    targetLang,
  }).pairs
}

/** Translate a batch of units in parallel; preserves order and per-unit status. */
export async function translateBatch(
  request: TranslateBatchRequest,
  opts: TranslateOneOptions,
): Promise<TranslateBatchResponse> {
  if (!Array.isArray(request.units) || request.units.length === 0) {
    return { ok: false, error: 'ai:translate-batch expected a non-empty `units` array' }
  }
  const memory = request.memoryEnabled === false
    ? null
    : (opts.memory ?? sharedMemory)
  const sourceLang = normalizeSourceLang(request.sourceLang)
  const targetLang = (request.targetLang ?? '').trim()
  const preserveFormat = request.preserveFormat !== false
  if (!targetLang) return { ok: false, error: 'ai:translate-batch expected non-empty `targetLang`' }

  const termPairs = terminologyForBatch(request, opts, sourceLang, targetLang)

  const settled = await Promise.all(
    request.units.map(async (unit): Promise<TranslateBatchUnitResult> => {
      const matchedTerms = matchTermsInSource(unit.sourceText, termPairs)
      const hit = memory?.lookup(sourceLang, targetLang, unit.sourceText)
      if (hit) {
        return {
          unitId: unit.unitId,
          sourceText: unit.sourceText,
          translatedText: hit.translatedText,
          status: 'memory-hit',
          range: unit.range ?? null,
          ...(matchedTerms.length > 0 ? { matchedTerms } : {}),
        }
      }
      const res = await translateOne(
        {
          instruction: unit.sourceText,
          sourceLang,
          targetLang,
          preserveFormat,
          range: unit.range ?? null,
          memoryEnabled: request.memoryEnabled,
          qualityCheck: request.qualityCheck,
          glossaryCategory: request.glossaryCategory,
        },
        opts,
      )
      const status: TranslateBatchUnitResult['status'] = res.ok ? 'translated' : 'failed'
      const warnings = res.ok ? warningsFor(unit, res.translated) : ['provider-error']
      const result: TranslateBatchUnitResult = {
        unitId: unit.unitId,
        sourceText: unit.sourceText,
        status,
        warnings,
        range: unit.range ?? null,
      }
      if (res.translated !== undefined) result.translatedText = res.translated
      if (!res.ok && res.error) result.errorMessage = res.error
      if (res.matchedTerms && res.matchedTerms.length > 0) result.matchedTerms = res.matchedTerms
      return result
    }),
  )

  const ok = settled.every((u) => u.status === 'translated' || u.status === 'memory-hit')
  const quality = assessBatchQuality(settled)
  const failed = settled.find((u) => u.status === 'failed')
  const response: TranslateBatchResponse = { ok, units: settled, quality }
  if (failed?.errorMessage) response.error = failed.errorMessage
  return response
}


/**
 * Options for {@link translateBatchStream}. The streaming variant is wire-shape
 * compatible with {@link translateBatch} but exposes per-unit callbacks so SSE
 * handlers can emit events as each unit settles. Concurrency is bounded so a
 * 500-unit document does not flood the LLM provider with simultaneous requests.
 */
export interface TranslateBatchStreamOptions {
  /**
   * Max in-flight provider calls at any moment. Defaults to 25 (matches the
   * Dataflare parent's chunk size). Lower values throttle provider load; higher
   * values increase throughput but risk rate-limits.
   */
  concurrency?: number
  /**
   * Called once per settled unit. Receives the unit index + total + final result.
   * The callback runs sequentially in completion order (not input order); SSE
   * handlers are expected to look up the unit by `unitId` if they need ordering.
   */
  onUnit?: (event: {
    index: number
    total: number
    result: TranslateBatchUnitResult
  }) => void | Promise<void>
}

/**
 * Streaming variant of {@link translateBatch}: settles units under a bounded
 * concurrency budget and fires `onUnit` for each completed unit. Returns the
 * same {@link TranslateBatchResponse} as the non-streaming variant, suitable for
 * callers that just need the aggregate (e.g. the IPC path that returns the
 * final shape to Electron docs / sheets / slides).
 */
export async function translateBatchStream(
  request: TranslateBatchRequest,
  opts: TranslateOneOptions,
  streamOpts: TranslateBatchStreamOptions = {},
): Promise<TranslateBatchResponse> {
  if (!Array.isArray(request.units) || request.units.length === 0) {
    return { ok: false, error: 'ai:translate-batch expected a non-empty `units` array' }
  }
  const memory = request.memoryEnabled === false
    ? null
    : (opts.memory ?? sharedMemory)
  const sourceLang = normalizeSourceLang(request.sourceLang)
  const targetLang = (request.targetLang ?? '').trim()
  const preserveFormat = request.preserveFormat !== false
  if (!targetLang) return { ok: false, error: 'ai:translate-batch expected non-empty `targetLang`' }

  const total = request.units.length
  const concurrency = Math.max(1, Math.min(streamOpts.concurrency ?? 25, total))
  const settled: TranslateBatchUnitResult[] = new Array(total)
  const termPairs = terminologyForBatch(request, opts, sourceLang, targetLang)

  // Build a unit-settler that re-uses the same per-unit logic as translateBatch.
  const settleOne = async (index: number): Promise<void> => {
    const unit = request.units[index]
    if (!unit) return
    const matchedTerms = matchTermsInSource(unit.sourceText, termPairs)
    const hit = memory?.lookup(sourceLang, targetLang, unit.sourceText)
    let result: TranslateBatchUnitResult
    if (hit) {
      result = {
        unitId: unit.unitId,
        sourceText: unit.sourceText,
        translatedText: hit.translatedText,
        status: 'memory-hit',
        range: unit.range ?? null,
        ...(matchedTerms.length > 0 ? { matchedTerms } : {}),
      }
    } else {
      const res = await translateOne(
        {
          instruction: unit.sourceText,
          sourceLang,
          targetLang,
          preserveFormat,
          range: unit.range ?? null,
          memoryEnabled: request.memoryEnabled,
          qualityCheck: request.qualityCheck,
          glossaryCategory: request.glossaryCategory,
        },
        opts,
      )
      const status: TranslateBatchUnitResult['status'] = res.ok ? 'translated' : 'failed'
      const warnings = res.ok ? warningsFor(unit, res.translated) : ['provider-error']
      result = {
        unitId: unit.unitId,
        sourceText: unit.sourceText,
        status,
        warnings,
        range: unit.range ?? null,
      }
      if (res.translated !== undefined) result.translatedText = res.translated
      if (!res.ok && res.error) result.errorMessage = res.error
      if (res.matchedTerms && res.matchedTerms.length > 0) result.matchedTerms = res.matchedTerms
    }
    settled[index] = result
    if (streamOpts.onUnit) {
      await streamOpts.onUnit({ index, total, result })
    }
  }

  // Bounded-concurrency driver: pull the next pending index when a slot frees up.
  let nextIndex = 0
  const workers = Array.from({ length: concurrency }, async () => {
    while (true) {
      const i = nextIndex++
      if (i >= total) return
      await settleOne(i)
    }
  })
  await Promise.all(workers)

  const ok = settled.every((u) => u.status === 'translated' || u.status === 'memory-hit')
  const quality = assessBatchQuality(settled)
  const failed = settled.find((u) => u.status === 'failed')
  const response: TranslateBatchResponse = { ok, units: settled, quality }
  if (failed?.errorMessage) response.error = failed.errorMessage
  return response
}
