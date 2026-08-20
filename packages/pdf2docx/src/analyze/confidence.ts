/**
 * Per-page conversion confidence (P4, plan §4.3): aggregates the pipeline's
 * quality signals into one 0..1 score. Pages under PAGE_CONFIDENCE_MIN are
 * downgraded by the caller to a full-page bitmap (pdf2htmlEX-style fidelity
 * floor) instead of shipping a garbled layout.
 */

/** below this aggregated confidence the page falls back to its bitmap */
export const PAGE_CONFIDENCE_MIN = 0.5

/** weight of the bad-ToUnicode ratio (0.15 ratio — the hard gate — ≈ −0.45) */
const BAD_UNICODE_WEIGHT = 3
/** cap of the bad-unicode penalty */
const BAD_UNICODE_PENALTY_MAX = 0.45
/** weight of the least-confident stream table (conf 0.5 table → −0.25) */
const STREAM_TABLE_WEIGHT = 0.5
/** penalty per layout warning (overlapping blocks, clamped gaps, …) */
const WARNING_PENALTY = 0.06
/** cap of the warning penalty */
const WARNING_PENALTY_MAX = 0.3

export interface ConfidenceSignals {
  /** share of U+FFFD / private-use code points among the page's text chars */
  badUnicodeRatio: number
  /** confidences of the page's stream (borderless) tables */
  streamTableConfidences: readonly number[]
  /** analysis-layer warning count for the page */
  warningCount: number
}

/** Aggregate the page signals into one confidence score (0..1). */
export function pageConfidence(s: ConfidenceSignals): number {
  let conf = 1
  conf -= Math.min(BAD_UNICODE_PENALTY_MAX, Math.max(0, s.badUnicodeRatio) * BAD_UNICODE_WEIGHT)
  if (s.streamTableConfidences.length > 0) {
    const weakest = Math.min(...s.streamTableConfidences)
    conf -= STREAM_TABLE_WEIGHT * Math.max(0, 1 - weakest)
  }
  conf -= Math.min(WARNING_PENALTY_MAX, WARNING_PENALTY * Math.max(0, s.warningCount))
  return Math.max(0, Math.min(1, conf))
}
