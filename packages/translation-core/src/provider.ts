import { type AiProviderConfig, type AiProviderId } from '@genoffice/ai-provider'

import {
  buildTranslationPrompt,
  buildTranslateSystemPrompt,
  extractTranslationText,
  normalizeSourceLang,
} from './prompt'
import { assessBatchQuality, warningsFor } from './quality'
import { TranslationMemory } from './memory'
import type {
  TranslateBatchRequest,
  TranslateBatchResponse,
  TranslateBatchUnitResult,
  TranslateRequest,
  TranslateResponse,
} from './types'
import { callLlm } from './llm-client'  // W9: seam between translation-core and the underlying LLM SDK

/**
 * `translateOne` / `translateBatch` are the public call surface shared by the
 * Electron main-process handlers and the web-server. They wrap
 * `chatForProvider` with the hardened prompt and quality checks so every call
 * site gets the same behaviour and the same error mapping.
 *
 * The shared instance is a singleton so consecutive calls (the typical
 * retranslation flow) hit the in-memory TM.
 */
export const sharedMemory = new TranslationMemory()

export interface TranslateOneOptions {
  provider: AiProviderId
  config: AiProviderConfig
  /** Optional external memory; falls back to the shared in-memory TM. */
  memory?: TranslationMemory
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
    }
  }

  const system = buildTranslateSystemPrompt({
    sourceLang,
    targetLang,
    preserveFormat,
    glossaryCategory: request.glossaryCategory,
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
  const translated = extractTranslationText(result.content ?? '')
  if (!translated) {
    return { ok: false, error: 'Translation response did not contain final text.' }
  }
  if (memory) memory.save({ sourceLang, targetLang, sourceText, translatedText: translated })
  return {
    ok: true,
    translated,
    planId: `translate-${Date.now().toString(36)}`,
    sourceLang,
    targetLang,
    preserveFormat,
    status: 'translated',
  }
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

  const settled = await Promise.all(
    request.units.map(async (unit): Promise<TranslateBatchUnitResult> => {
      const hit = memory?.lookup(sourceLang, targetLang, unit.sourceText)
      if (hit) {
        return {
          unitId: unit.unitId,
          sourceText: unit.sourceText,
          translatedText: hit.translatedText,
          status: 'memory-hit',
          range: unit.range ?? null,
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

  // Build a unit-settler that re-uses the same per-unit logic as translateBatch.
  const settleOne = async (index: number): Promise<void> => {
    const unit = request.units[index]
    const hit = memory?.lookup(sourceLang, targetLang, unit.sourceText)
    let result: TranslateBatchUnitResult
    if (hit) {
      result = {
        unitId: unit.unitId,
        sourceText: unit.sourceText,
        translatedText: hit.translatedText,
        status: 'memory-hit',
        range: unit.range ?? null,
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
