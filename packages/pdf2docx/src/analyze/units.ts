/**
 * Line units — a visual line split at column-scale x gaps. The shared
 * substrate for stream-table detection and XY-Cut column detection: both need
 * sub-line granularity because PDFium's content order can merge text of
 * side-by-side columns / table cells into one visual line.
 */
import type { Rect } from '../geometry'
import { median, rectUnion } from '../geometry'
import type { PdfChar } from '../ir'
import type { RawLine } from './lines'
import { isSpaceCode } from './lines'
import { groupIntoWords, isLetterSpacedLine, medianCharGap, spaceGapThreshold } from './words'

/** a gap this many times the word-space threshold splits a line into units… */
const COLUMN_GAP_SPACE_MULT = 2
/** …and never narrower than this many ems (column gaps are visibly wide) */
const COLUMN_GAP_MIN_EMS = 0.9

/** units on the same visual row cluster within this share of the font size */
const ROW_BASELINE_TOL_RATIO = 0.45

export interface LineUnit {
  /** chars in x order; interior whitespace kept (word boundaries), edges trimmed */
  chars: PdfChar[]
  box: Rect
  baseline: number
  /** index of the source RawLine; units of one visual line share it */
  lineIndex: number
  /** word count (CJK counts one per char) — the "sentence-ness" signal */
  wordCount: number
  /** median font size of the unit's chars */
  fontSize: number
}

const isVisible = (c: PdfChar): boolean => !isSpaceCode(c.code) && c.code > 0x1f

function makeUnit(chars: PdfChar[], lineIndex: number): LineUnit {
  const visible = chars.filter(isVisible)
  const box = visible.map((c) => c.box).reduce(rectUnion)
  return {
    chars,
    box,
    baseline: median(visible.map((c) => c.originY)),
    lineIndex,
    wordCount: groupIntoWords(chars).length,
    fontSize: median(visible.map((c) => c.fontSize)) || 12,
  }
}

/** the x gap (pt) that separates cell/column content within one line */
export function columnGapThreshold(fontSize: number, lineGap: number): number {
  return Math.max(
    COLUMN_GAP_SPACE_MULT * spaceGapThreshold(fontSize, lineGap),
    COLUMN_GAP_MIN_EMS * fontSize,
  )
}

/**
 * Split raw lines into units at column-scale gaps. Chars are x-sorted first
 * (RTL lines arrive in visual/content order; geometry is what matters here).
 */
export function splitIntoUnits(rawLines: readonly RawLine[]): LineUnit[] {
  const units: LineUnit[] = []
  for (const [lineIndex, raw] of rawLines.entries()) {
    const sorted = [...raw.chars].sort((a, b) => a.looseBox.x0 - b.looseBox.x0)
    const visible = sorted.filter(isVisible)
    if (visible.length === 0) continue
    const fontSize = median(visible.map((c) => c.fontSize)) || 12
    // letter-spaced display text (P10 C): uniform tracking gaps are one line
    // of one run, not columns — splitting them stacks each glyph as its own
    // block and the title collapses into an overlapped heap
    const threshold = isLetterSpacedLine(sorted)
      ? Infinity
      : columnGapThreshold(fontSize, medianCharGap(visible))
    let current: PdfChar[] = []
    let lastVisible: PdfChar | null = null
    for (const c of sorted) {
      if (!isVisible(c)) {
        // interior whitespace stays with its unit; leading whitespace drops.
        // Displaced whitespace drops too (P20): Word-export PDFs park a
        // trailing NBSP at the line's LEFT edge, and the x-sort above lands
        // it inside the first word ("Contact" → "C ontact") — a space glyph
        // overlapping the previous visible glyph separates nothing.
        const cx = (c.box.x0 + c.box.x1) / 2
        if (current.length > 0 && (lastVisible === null || cx > lastVisible.box.x1 - 0.1))
          current.push(c)
        continue
      }
      if (lastVisible && c.looseBox.x0 - lastVisible.looseBox.x1 > threshold) {
        units.push(makeUnit(current, lineIndex))
        current = []
      }
      current.push(c)
      lastVisible = c
    }
    if (current.length > 0) units.push(makeUnit(current, lineIndex))
  }
  return units
}

export interface UnitRow {
  units: LineUnit[]
  box: Rect
  baseline: number
}

/**
 * Cluster units into visual rows (top → bottom). Units of one RawLine always
 * share a row; separately-extracted cells on the same baseline join it too.
 */
export function clusterUnitRows(units: readonly LineUnit[]): UnitRow[] {
  const sorted = [...units].sort((a, b) => b.baseline - a.baseline)
  const rows: UnitRow[] = []
  for (const unit of sorted) {
    const last = rows[rows.length - 1]
    const tol = ROW_BASELINE_TOL_RATIO * Math.max(unit.fontSize, 1)
    if (last && Math.abs(unit.baseline - last.baseline) <= tol) {
      last.units.push(unit)
      last.box = rectUnion(last.box, unit.box)
    } else {
      rows.push({ units: [unit], box: unit.box, baseline: unit.baseline })
    }
  }
  for (const row of rows) row.units.sort((a, b) => a.box.x0 - b.box.x0)
  return rows
}
