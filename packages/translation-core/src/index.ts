/**
 * Public surface of the AI translation core.
 *
 *   import {
 *     translateOne, translateBatch, translateBatchStream,
 *     sharedMemory,
 *     chunkDocument, buildTranslationPrompt, extractTranslationText,
 *     LANGUAGES, englishLabelFor,
 *     assessQuality, assessBatchQuality,
 *   } from '@genoffice/translation-core'
 *
 * The Electron main-process handlers (docs / sheets / slides) and the
 * web-server `ai:translate` route all funnel through `translateOne` and
 * `translateBatch`, so a prompt tweak or quality-rule change in this package
 * reaches every call site in one step.
 *
 * `translateBatchStream` is the per-unit callback variant used by the
 * standalone web-server's `/api/ai/translate/stream` SSE endpoint and the
 * Dataflare parent's bridge forwarder; it shares logic with `translateBatch`
 * but settles units under a bounded concurrency budget.
 */

export type {
  EditorRange,
  LanguageCode,
  LanguageOption,
  QualityReport,
  TranslateBatchRequest,
  TranslateBatchResponse,
  TranslateBatchUnitResult,
  TranslateRequest,
  TranslateResponse,
  TranslationUnit,
} from './types'

export { LANGUAGES, getLanguage, englishLabelFor } from './languages'

export {
  buildTranslationPrompt,
  buildTranslateSystemPrompt,
  extractTranslationText,
  normalizeSourceLang,
} from './prompt'

export { chunkDocument, makeUnitId } from './chunking'

export { assessQuality, assessBatchQuality, warningsFor } from './quality'

export {
  TranslationMemory,
  unitMemoryKey,
  type MemoryEntry,
  type MemorySaveRequest,
  type MemorySaveResponse,
} from './memory'

export { translateOne, translateBatch, translateBatchStream, sharedMemory } from './provider'
export type { TranslateOneOptions, TranslateBatchStreamOptions } from './provider'

// W9 seam — the LLM client boundary. Hosts may swap callers via setLlmCaller().
export {
  callLlm,
  callLlmWith,
  setLlmCaller,
  getLlmCaller,
  aiProviderCaller,
  piAiCaller,
  type LlmCaller,
  type LlmCallOptions,
  type LlmCallResult,
} from './llm-client'
