import {
  chatForProvider,
  isAiOverloadedError,
  type AiProviderConfig,
  type AiProviderId,
} from '@genoffice/ai-provider'

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
  const memory = opts.memory ?? sharedMemory
  const sourceLang = normalizeSourceLang(request.sourceLang)
  const preserveFormat = request.preserveFormat !== false

  const hit = memory.lookup(sourceLang, targetLang, sourceText)
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
  })
  try {
    const result = await chatForProvider(
      opts.provider,
      opts.config,
      system,
      buildTranslationPrompt(sourceText),
    )
    if (!result.ok) {
      return {
        ok: false,
        error: isAiOverloadedError(result.error)
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
    memory.save({ sourceLang, targetLang, sourceText, translatedText: translated })
    return {
      ok: true,
      translated,
      planId: `translate-${Date.now().toString(36)}`,
      sourceLang,
      targetLang,
      preserveFormat,
      status: 'translated',
    }
  } catch (err) {
    return {
      ok: false,
      error: isAiOverloadedError(err)
        ? 'AI service is busy — please retry shortly.'
        : err instanceof Error
          ? err.message
          : String(err),
    }
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
  const memory = opts.memory ?? sharedMemory
  const sourceLang = normalizeSourceLang(request.sourceLang)
  const targetLang = (request.targetLang ?? '').trim()
  const preserveFormat = request.preserveFormat !== false
  if (!targetLang) return { ok: false, error: 'ai:translate-batch expected non-empty `targetLang`' }

  const settled = await Promise.all(
    request.units.map(async (unit): Promise<TranslateBatchUnitResult> => {
      const hit = memory.lookup(sourceLang, targetLang, unit.sourceText)
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
