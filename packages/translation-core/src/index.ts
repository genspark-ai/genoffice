/**
 * Public surface of the AI translation core.
 *
 *   import {
 *     translateOne, translateBatch,
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

export { translateOne, translateBatch, sharedMemory } from './provider'
export type { TranslateOneOptions } from './provider'
