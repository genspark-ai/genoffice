/**
 * Public types for the AI translation core.
 *
 * Three call sites share these shapes:
 *  1. the Electron main-process IPC handlers in docs / sheets / slides
 *  2. the standalone web-server (`apps/web-server/src/ai/chat.ts`)
 *  3. the Dataflare bridge in the docs renderer (`apps/docs/src/renderer/web-bridge.ts`)
 *
 * Keep these in lock-step with the renderer-side TypeScript interfaces declared
 * in `apps/docs/src/shared/ipc.ts` (the `aiTranslate` / `aiTranslateBatch`
 * request and response shapes) — those are the wire contract.
 *
 * Note: optional fields use `T | undefined` (not bare `T?`) so the types
 * survive `exactOptionalPropertyTypes: true` callers in the sheets / docs
 * Electron builds.
 */

export type LanguageCode =
  | 'auto'
  | 'zh-CN'
  | 'zh-TW'
  | 'en-US'
  | 'ja-JP'
  | 'ko-KR'
  | 'fr-FR'
  | 'de-DE'
  | 'es-ES'
  | 'it-IT'
  | 'pt-PT'
  | 'ru-RU'
  | 'ar-SA'
  | 'hi-IN'
  | 'th-TH'

export interface LanguageOption {
  value: LanguageCode
  label: string
  englishLabel: string
}

export interface EditorRange {
  from?: number | undefined
  to?: number | undefined
  scope?: 'selection' | 'document' | 'paragraph' | 'cell' | 'table' | undefined
}

export interface TranslationUnit {
  unitId: string
  kind: 'paragraph' | 'heading' | 'list-item' | 'table-cell' | 'document'
  sourceText: string
  order: number
  range?: EditorRange | null | undefined
  metadata?: Record<string, unknown> | undefined
}

export interface TranslateRequest {
  instruction: string
  sourceLang?: LanguageCode | string | undefined
  targetLang: LanguageCode | string
  preserveFormat?: boolean | undefined
  range?: EditorRange | null | undefined
}

export interface TranslateResponse {
  ok: boolean
  translated?: string | undefined
  planId?: string | undefined
  error?: string | undefined
  sourceLang?: string | undefined
  targetLang?: string | undefined
  preserveFormat?: boolean | undefined
  status?: 'translated' | 'memory-hit' | 'failed' | undefined
  matchedTerms?: string[] | undefined
  warnings?: string[] | undefined
}

export interface TranslateBatchRequest {
  units: TranslationUnit[]
  sourceLang?: LanguageCode | string | undefined
  targetLang: LanguageCode | string
  preserveFormat?: boolean | undefined
  scene?: string | undefined
}

export interface TranslateBatchUnitResult {
  unitId: string
  sourceText: string
  translatedText?: string | undefined
  status?: 'translated' | 'memory-hit' | 'failed' | undefined
  matchedTerms?: string[] | undefined
  warnings?: string[] | undefined
  errorMessage?: string | undefined
  range?: EditorRange | null | undefined
}

export interface TranslateBatchResponse {
  ok: boolean
  units?: TranslateBatchUnitResult[] | undefined
  quality?: { overallScore?: number | undefined; warnings?: string[] | undefined } | undefined
  error?: string | undefined
}

export interface QualityReport {
  overallScore: number
  warnings: string[]
  empty?: boolean
  untranslated?: boolean
}
