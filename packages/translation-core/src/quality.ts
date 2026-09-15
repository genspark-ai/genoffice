import type { QualityReport, TranslationUnit } from './types'

/**
 * Cheap, deterministic quality checks run after each translation.
 *
 * The LLM is told to translate; if the answer is empty, suspiciously short,
 * suspiciously close to a verbatim copy of the source for a non-trivial input,
 * or repeats the same boilerplate that signals a refusal, the renderer flags
 * the unit so the user can retry. These are heuristics — they catch the most
 * common failures without needing another model call.
 */
export function assessQuality(
  sourceText: string,
  translatedText: string | null | undefined,
): QualityReport {
  const trimmed = (translatedText ?? '').trim()
  const src = (sourceText ?? '').trim()
  const warnings: string[] = []
  if (!trimmed) {
    return { overallScore: 0, warnings: ['empty'], empty: true }
  }
  if (REFUSAL_MARKERS.some((m) => trimmed.toLowerCase().includes(m))) {
    warnings.push('refusal-marker')
  }
  if (src.length > 0) {
    const ratio = trimmed.length / src.length
    if (ratio < MIN_RATIO) warnings.push('too-short')
    if (ratio > MAX_RATIO) warnings.push('too-long')
    // verbatim copy on a meaningful input → model probably echoed source
    if (src.length >= MIN_UNTRANSLATED_LENGTH && trimmed.toLowerCase() === src.toLowerCase()) {
      warnings.push('untranslated')
    }
  }
  const score = warnings.length === 0 ? 1 : Math.max(0, 1 - warnings.length * 0.25)
  return {
    overallScore: score,
    warnings,
    untranslated: warnings.includes('untranslated'),
    empty: false,
  }
}

const MIN_RATIO = 0.25
const MAX_RATIO = 6.0
const MIN_UNTRANSLATED_LENGTH = 12

const REFUSAL_MARKERS = [
  'i cannot translate',
  "i can't translate",
  'i am unable to translate',
  "i'm unable to translate",
  'as an ai',
  'i cannot help',
  'sorry, i cannot',
  '对不起，我无法翻译',
  '抱歉，我无法翻译',
]

/** Aggregate unit reports into a single batch-level quality report. */
export function assessBatchQuality(
  units: ReadonlyArray<{
    sourceText: string
    translatedText?: string | undefined
    warnings?: string[] | undefined
  }>,
): { overallScore: number; warnings: string[] } {
  if (units.length === 0) return { overallScore: 1, warnings: [] }
  const reports: QualityReport[] = units.map((u) => {
    const r = assessQuality(u.sourceText, u.translatedText)
    return { ...r, warnings: Array.from(new Set([...(r.warnings ?? []), ...(u.warnings ?? [])])) }
  })
  const score = reports.reduce((s, r) => s + r.overallScore, 0) / reports.length
  const warnings = Array.from(new Set(reports.flatMap((r) => r.warnings)))
  return { overallScore: score, warnings }
}

/** Tiny helper used by `provider.ts` to surface per-unit warnings in the response. */
export function warningsFor(
  unit: TranslationUnit,
  translated: string | null | undefined,
): string[] {
  return assessQuality(unit.sourceText, translated).warnings
}
