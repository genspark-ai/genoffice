import { canvasMetrics, type FontMetricsProvider } from '../line-metrics'
import { maxLineWidthPx, maxWordWidthPx, type MeasureRun } from './tokenize'

export interface RunDefaults {
  sizeHalfPoints?: number
  font?: string
  /** rPrDefault w:lang w:eastAsia: the break rules beside a middle dot follow it */
  eastAsiaLang?: string
}

export type { MeasureRun }

/**
 * Text measurement seam of the column-width algorithm: min-content (widest
 * unbreakable token) and max-content (unwrapped line) of a cell paragraph.
 * `measured` tells the algorithm whether the advances are real font metrics
 * (no slack, a px edge per side) or the heuristic estimate (8 % slack).
 */
export interface TextMeasure {
  wordPx(runs: MeasureRun[]): number
  linePx(runs: MeasureRun[]): number
  measured: boolean
}

const heuristicMeasure: TextMeasure = {
  wordPx: (runs) => maxWordWidthPx(runs),
  linePx: (runs) => maxLineWidthPx(runs),
  measured: false,
}

export function textMeasureOf(metrics: FontMetricsProvider | null): TextMeasure {
  if (!metrics) return heuristicMeasure
  return {
    wordPx: (runs) => maxWordWidthPx(runs, metrics),
    linePx: (runs) => maxLineWidthPx(runs, metrics),
    measured: true,
  }
}

/** the live app measures with canvas advances; node / jsdom falls back to the heuristic */
export function defaultTextMeasure(): TextMeasure {
  return textMeasureOf(canvasMetrics())
}
