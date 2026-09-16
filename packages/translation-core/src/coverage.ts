/**
 * Translation coverage — how much of a file the dictionary will actually reach.
 *
 * The format handlers are pure dictionary rewriters: `translate_ppt.py`,
 * `translate_docx.py`, `translate_pdf.py` and `translate_xls.py` all walk the
 * file, and for each text unit either replace a dictionary key (exact match
 * first, then longest-key-first substring replacement) or leave the unit
 * exactly as it was. There is no model call inside the file pass.
 *
 * That makes "did this file actually get translated?" a question about the
 * dictionary, not about the translator — and one the user has to be able to
 * answer *before* trusting the output. A deck whose dictionary covers 8 of 40
 * paragraphs comes back looking translated while 32 paragraphs are silently
 * still in the source language.
 *
 * The handlers' own untranslated report cannot answer it either: docx, xls and
 * pdf only flag a missed string when it contains Latin letters or kana (a
 * heuristic tuned for `targetLang: zh-CN`), so a zh → en pass reports zero
 * misses while translating nothing. Recomputing the coverage here — from the
 * segments we mined and the dictionary we wrote — gives the same answer for
 * every format and every language pair.
 *
 * This module is deliberately dependency-free so both the dictionary builder
 * and the host can use it.
 */

/** Everything the caller needs to explain a coverage number to a user. */
export interface CoverageReport {
  /** Segments mined from the file (after `minChars`/dedupe filtering). */
  total: number
  /** Segments at least one dictionary key will change (`exact + partial`). */
  covered: number
  /** Segments the dictionary rewrites in full (the segment *is* a key). */
  exact: number
  /**
   * Segments the dictionary only rewrites in part — a key appears inside a
   * longer string. These come out mixed-language ("Product Acceptance
   * Report已提交。"), which reads worse than leaving them alone, so they are
   * reported separately instead of being folded into `covered`.
   */
  partial: string[]
  /** Segments that will come out in the source language, first-seen order. */
  uncovered: string[]
  /** `covered / total`, or 1 when there is nothing to translate. */
  ratio: number
}

/**
 * A dictionary of `source -> target` pairs. Values may be empty strings for a
 * template the user has not filled in yet; an empty target does not count as
 * coverage.
 */
export type TranslationDictionary = Readonly<Record<string, string | undefined>>

/**
 * True when the handlers would change `segment`.
 *
 * Mirrors their lookup exactly: an exact key hit short-circuits, otherwise the
 * longest key that appears anywhere in the segment is substituted. An empty
 * target is not coverage — the handler returns the segment untouched
 * (`tgt if tgt is not None else result` on the exact path, and the substring
 * loop skips `tgt is None`), and a blank string would delete the text.
 */
export function isSegmentCovered(segment: string, dictionary: TranslationDictionary): boolean {
  if (!segment) return false
  const exact = dictionary[segment]
  if (typeof exact === 'string' && exact.trim().length > 0) return true
  for (const [source, target] of Object.entries(dictionary)) {
    if (!source) continue
    if (typeof target !== 'string' || target.trim().length === 0) continue
    if (segment.includes(source)) return true
  }
  return false
}

/**
 * Split `segments` into the ones the dictionary reaches and the ones it does
 * not. Order is preserved so the first rows of a table are reported first.
 */
export function assessCoverage(
  segments: readonly string[],
  dictionary: TranslationDictionary,
): CoverageReport {
  const uncovered: string[] = []
  const partial: string[] = []
  let covered = 0
  let exact = 0
  for (const segment of segments) {
    if (!isSegmentCovered(segment, dictionary)) {
      uncovered.push(segment)
      continue
    }
    covered++
    const hit = dictionary[segment]
    if (typeof hit === 'string' && hit.trim().length > 0) exact++
    else partial.push(segment)
  }
  const total = segments.length
  return { total, covered, exact, partial, uncovered, ratio: total === 0 ? 1 : covered / total }
}

/**
 * Merge new pairs into an existing dictionary.
 *
 * Existing keys win: the user may have hand-edited a generated dictionary, and
 * a later gap-filling pass must not clobber that. Returns a new object sorted
 * by key so the file stays diffable.
 */
export function mergeDictionary(
  existing: TranslationDictionary,
  added: TranslationDictionary,
): Record<string, string> {
  const merged: Record<string, string> = {}
  for (const [source, target] of Object.entries(existing)) {
    if (typeof target !== 'string') continue
    if (!source) continue
    merged[source] = target
  }
  for (const [source, target] of Object.entries(added)) {
    if (!source || typeof target !== 'string') continue
    if (merged[source] !== undefined) continue
    merged[source] = target
  }
  const sorted: Record<string, string> = {}
  for (const key of Object.keys(merged).sort()) sorted[key] = merged[key] as string
  return sorted
}
