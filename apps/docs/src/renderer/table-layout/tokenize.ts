import {
  DEFAULT_FONT_FAMILY,
  DEFAULT_FONT_SIZE_PT,
  HeuristicMetrics,
  isCjk,
  textHasHangul,
  type FontMetricsProvider,
  type RunStyle,
} from '../line-metrics'
import { cjkForbidsLineEnd, cjkForbidsLineStart } from './cjk-break'

/** one run of a cell paragraph as the tokenizer sees it */
export interface MeasureRun {
  text: string
  fontFamily?: string
  sizeHalfPoints?: number
  bold?: boolean
  italic?: boolean
  eastAsiaLang?: string
}

const heuristicMetrics = new HeuristicMetrics()

const MIDDLE_DOT = '\u00b7'
const BREAK_AFTER_DASHES = new Set(['-', '\u2013', '\u2014'])

/**
 * Word's break rules beside a middle dot / before a colon follow the run's
 * East Asian language (probe 2026-11-08): zh-CN (and no w:lang
 * w:eastAsia at all, Word's own default) lets neither side of "\u00b7" break,
 * zh-TW only keeps the character off a line start, every other
 * language breaks at the spaces like anywhere else.
 */
export type MiddleDotGlue = 'both' | 'before' | 'none'

export function middleDotGlue(eastAsiaLang?: string): MiddleDotGlue {
  if (!eastAsiaLang) return 'both'
  const tag = eastAsiaLang.toLowerCase()
  if (!tag.startsWith('zh')) return 'none'
  return /-(tw|hk|mo|hant)\b/.test(tag) ? 'before' : 'both'
}

/**
 * Widest unbreakable token width (px) across runs: whitespace splits words;
 * CJK breaks per character unless Hangul is present (keep-all, same rule as
 * simulateLines). Min-content input for table autofit; scanning is capped so
 * huge cells stay cheap.
 */
export function maxWordWidthPx(
  runs: MeasureRun[],
  metrics: FontMetricsProvider = heuristicMetrics,
  maxWords = 200,
): number {
  const keepAll = runs.some((r) => textHasHangul(r.text))
  const full = runs.map((r) => r.text).join('')
  let pos = 0
  let max = 0
  let cur = 0 // width of the current token's completed segments (may span runs)
  let words = 0
  let scanned = 0
  let breakAfter = false // pending break opportunity behind a '-'
  let afterCjk = false // the token so far ends in an East Asian character
  let glueNext = false // ... which forbids a line end (opener)
  const endWord = () => {
    if (cur > max) max = cur
    if (cur > 0) words++
    cur = 0
    afterCjk = false
    glueNext = false
  }
  for (const run of runs) {
    if (!run.text) continue
    const style: RunStyle = {
      fontFamily: run.fontFamily ?? DEFAULT_FONT_FAMILY,
      fontSizePx: (run.sizeHalfPoints ? run.sizeHalfPoints / 2 : DEFAULT_FONT_SIZE_PT) * (96 / 72),
      bold: !!run.bold,
      italic: !!run.italic,
    }
    const glue = middleDotGlue(run.eastAsiaLang)
    let buf = ''
    const flushSeg = () => {
      if (!buf) return
      cur += metrics.measure(buf, style)
      buf = ''
    }
    for (const ch of run.text) {
      if (words >= maxWords || scanned >= maxWords * 60) {
        flushSeg()
        endWord()
        return max
      }
      scanned++
      // Word breaks after '-' even before a digit ("10-|20", "DDR5-|6400") and
      // after an en/em dash ("R7–|R9"); never at '/' or U+2212 ("mmmm/mmmm",
      // "3.2/R1", "01/02", "mmmm−mmmm" stay whole; probes 2026-11-08)
      if (breakAfter) {
        flushSeg()
        endWord()
        breakAfter = false
      }
      const glued =
        ch === ' ' &&
        glue !== 'none' &&
        (full[pos + 1] === MIDDLE_DOT ||
          full[pos + 1] === ':' ||
          (glue === 'both' && full[pos - 1] === MIDDLE_DOT))
      pos += ch.length
      if (!glued && (ch === ' ' || ch === '\t' || ch === '\n')) {
        flushSeg()
        endWord()
        continue
      }
      if (BREAK_AFTER_DASHES.has(ch)) breakAfter = true
      if (!keepAll && isCjk(ch.codePointAt(0) ?? 0)) {
        flushSeg()
        // any two East Asian characters may break apart, a closer stays with
        // its predecessor and an opener with its successor (probe 2026-11-09)
        if (!cjkForbidsLineStart(ch) && !glueNext) endWord()
        cur += metrics.measure(ch, style)
        afterCjk = true
        glueNext = cjkForbidsLineEnd(ch)
        continue
      }
      if (afterCjk) {
        // Latin beside East Asian text is its own token unless an opener holds it
        flushSeg()
        if (glueNext) afterCjk = glueNext = false
        else endWord()
      }
      buf += ch
    }
    flushSeg()
  }
  endWord()
  return max
}

/**
 * Unwrapped width (px) of the runs laid on one line; hard breaks ('\n') start
 * a new line and the widest wins. Max-content input for table autofit.
 */
export function maxLineWidthPx(
  runs: MeasureRun[],
  metrics: FontMetricsProvider = heuristicMetrics,
): number {
  let max = 0
  let cur = 0
  for (const run of runs) {
    if (!run.text) continue
    const style: RunStyle = {
      fontFamily: run.fontFamily ?? DEFAULT_FONT_FAMILY,
      fontSizePx: (run.sizeHalfPoints ? run.sizeHalfPoints / 2 : DEFAULT_FONT_SIZE_PT) * (96 / 72),
      bold: !!run.bold,
      italic: !!run.italic,
    }
    const lines = run.text.split('\n')
    lines.forEach((line, i) => {
      if (i > 0) {
        max = Math.max(max, cur)
        cur = 0
      }
      if (line) cur += metrics.measure(line, style)
    })
  }
  return Math.max(max, cur)
}
