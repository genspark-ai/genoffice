/**
 * Public surface of the AI translation core.
 *
 *   import {
 *     translateOne, translateBatch, translateBatchStream,
 *     sharedMemory,
 *     chunkDocument, buildTranslationPrompt, extractTranslationText,
 *     LANGUAGES, englishLabelFor,
 *     assessQuality, assessBatchQuality,
 *     KnowledgeBase, PersistentTranslationMemory,
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
 *
 * `KnowledgeBase` and `PersistentTranslationMemory` are the storage layer that
 * mirrors the LumosAI translation suite (`translate-config` 5-schema KB +
 * JSON-backed TM with fuzzy match). Hosts pass them into `translateOne` /
 * `translateBatch` via the `TranslateOneOptions.knowledgeBase` and
 * `TranslateOneOptions.memory` slots — both are optional, so existing call
 * sites keep working unchanged.
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

export { translateOne, translateBatch, translateBatchStream, sharedMemory,
  type TranslationMemoryLike } from './provider'
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

// W35+ translation knowledge base + persistent TM — mirrors LumosAI's
// translate-config / translation_memory layout under
// `~/.genoffice/translation-kb.json` and `~/.genoffice/translation-memory/`.
export {
  KnowledgeBase,
  SCOPES,
  SCHEMA_IDS,
  SCHEMA_KEY_TO_ID,
  SCHEMA_TO_KEY,
  type BrandEntry,
  type CustomerPreferenceEntry,
  type ForbiddenEntry,
  type KBEntry,
  type KBListFilters,
  type KBStore,
  type KnowledgeBaseFileSystem,
  type KnowledgeBaseOptions,
  type ResolvedRules,
  type Scope,
  type SchemaId,
  type SchemaKey,
  type StyleRuleEntry,
  type TermEntry,
} from './knowledge-base'

export {
  PersistentTranslationMemory,
  type PersistentLookupHit,
  type PersistentMemoryFileSystem,
  type PersistentTranslationMemoryOptions,
} from './persistent-memory'

// Whole-file translation bridge + KB/LLM dictionary builder. Both are
// host-agnostic (Node built-ins + `@genoffice/file-parse` only) so the Electron
// main process and the standalone web-server share one implementation.
export {
  SUPPORTED_EXTENSIONS,
  defaultOutputPath,
  isSupportedExtension,
  resolveTranslateSkills,
  translateFile,
  type SupportedExtension,
  type TranslateFileRequest,
  type TranslateFileResult,
  type TranslateSkillsLocation,
} from './file-translate'

// Terminology enforcement for the live translate path (KB terms + generated
// dictionary). `dictionary.ts` uses the same rules while *building* a
// dictionary; this module owns them for a single selection / snippet.
export {
  applyTerminology,
  matchTermsInSource,
  resolveKbForCall,
  terminologyPairs,
  type ResolveKbOptions,
  type TerminologyPair,
} from './kb-rules'

export {
  applyKbRules,
  batchSegments,
  buildDictionary,
  defaultDictionaryPath,
  mineSegments,
  type BuildDictionaryDeps,
  type BuildDictionaryRequest,
  type BuildDictionaryResult,
  type DictionarySegment,
} from './dictionary'
