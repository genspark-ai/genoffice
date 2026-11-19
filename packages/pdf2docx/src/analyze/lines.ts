/**
 * char → line grouping (pure geometry, ported from pdftext's aggregation
 * rules). Input chars arrive in PDFium content order, which is reading order
 * for the single-column documents P1 targets.
 */
import type { Rect } from '../geometry'
import { median, rectArea, rectUnion, verticalOverlapRatio } from '../geometry'
import type { PdfChar } from '../ir'
import { isCombiningMark } from '../script'

export interface RawLine {
  /** visible chars in order (whitespace kept for the word layer) */
  chars: PdfChar[]
  baseline: number
  box: Rect
  endsWithHyphen: boolean
}

/** baseline tolerance as a share of font size (same-line clustering) */
const BASELINE_TOL_RATIO = 0.45
/** chars whose origin regresses further than this (in ems) start a new line */
const X_REGRESSION_EMS = 1.5
/** PDFium's generated hyphenation marker code point (see pdftext) */
const HYPHEN_MARKER = 0x02

const isNewlineCode = (code: number): boolean => code === 0x0a || code === 0x0d
export const isSpaceCode = (code: number): boolean =>
  code === 0x20 || code === 0xa0 || code === 0x09 || code === 0x3000

const isVisible = (c: PdfChar): boolean => !isSpaceCode(c.code) && c.code > 0x1f

/**
 * Merge combining marks / zero-width code points into the preceding base
 * character's cluster so downstream layers never split them off.
 */
export function clusterCombiningMarks(chars: readonly PdfChar[]): PdfChar[] {
  const out: PdfChar[] = []
  for (const c of chars) {
    const prev = out[out.length - 1]
    if (prev && isCombiningMark(c.code) && isVisible(prev)) {
      out[out.length - 1] = {
        ...prev,
        text: prev.text + c.text,
        box: rectArea(c.box) > 0 ? rectUnion(prev.box, c.box) : prev.box,
      }
      continue
    }
    out.push(c)
  }
  return out
}

interface OpenLine {
  chars: PdfChar[]
  /** running baseline (median-updated at finalize; building compares locally) */
  baseline: number
  box: Rect | null
  lastVisible: PdfChar | null
}

function finalizeLine(open: OpenLine): RawLine | null {
  const visible = open.chars.filter(isVisible)
  if (visible.length === 0) return null
  const box = visible.map((c) => c.box).reduce(rectUnion)
  // the flag must survive PDFium's 0x02 hyphen-replacement char, which
  // isVisible filters as a control code — scan the tail past it (P21 B)
  let last: PdfChar | undefined
  for (let i = open.chars.length - 1; i >= 0; i--) {
    const c = open.chars[i]!
    if (isVisible(c) || c.code === HYPHEN_MARKER || c.isHyphen) {
      last = c
      break
    }
  }
  return {
    chars: open.chars,
    baseline: median(visible.map((c) => c.originY)),
    box,
    endsWithHyphen: last !== undefined && (last.isHyphen || last.code === HYPHEN_MARKER),
  }
}

/** the char's baseline continues the open line's (tolerance-based) */
function onBaseline(open: OpenLine, c: PdfChar): boolean {
  const anchor = open.lastVisible
  if (!anchor) return true
  const baselineTol = BASELINE_TOL_RATIO * Math.max(c.fontSize, anchor.fontSize, 1)
  return (
    Math.abs(c.originY - anchor.originY) <= baselineTol ||
    Math.abs(c.originY - open.baseline) <= baselineTol
  )
}

function sameLine(open: OpenLine, c: PdfChar): boolean {
  const anchor = open.lastVisible
  if (!anchor) return true
  // super/subscripts sit off the baseline but overlap the line box vertically
  const overlaps =
    open.box !== null && rectArea(c.box) > 0 && verticalOverlapRatio(c.box, open.box) >= 0.5
  if (!onBaseline(open, c) && !overlaps) return false
  // a hard x regression on the same baseline is a new visual line (or column)
  return c.originX >= anchor.originX - X_REGRESSION_EMS * Math.max(c.fontSize, anchor.fontSize, 1)
}

/** Group clustered chars into visual lines. Whitespace stays inside its line. */
export function groupIntoLines(chars: readonly PdfChar[]): RawLine[] {
  const lines: RawLine[] = []
  let open: OpenLine | null = null
  let pendingBreak = false
  let pendingGenerated = false

  const flush = (): void => {
    if (!open) return
    const line = finalizeLine(open)
    if (line) lines.push(line)
    open = null
  }

  for (const c of chars) {
    if (c.code <= 0) continue
    if (isNewlineCode(c.code)) {
      // PDFium fabricates newlines from x-gap heuristics, and letter-spaced
      // display text (every glyph its own text op) trips it mid-line (P10 C).
      // A generated break only holds when the next visible char really leaves
      // the baseline; real newline glyphs always break.
      if (c.isGenerated) pendingGenerated = true
      else pendingBreak = true
      continue
    }
    // leading whitespace after a break contributes nothing
    if (!open && isSpaceCode(c.code)) continue

    if (open) {
      if (isVisible(c)) {
        const generatedHolds = pendingGenerated && !onBaseline(open, c)
        if (pendingBreak || generatedHolds || !sameLine(open, c)) flush()
      } else if (pendingBreak) {
        flush()
      }
    }
    if (isVisible(c) || !open) {
      pendingBreak = false
      pendingGenerated = false
    }

    open ??= { chars: [], baseline: c.originY, box: null, lastVisible: null }
    open.chars.push(c)
    if (isVisible(c)) {
      open.box = open.box ? rectUnion(open.box, c.box) : c.box
      open.lastVisible = c
      // drift slowly toward the latest baseline so long lines with slight
      // slope keep clustering; finalize recomputes the true median
      open.baseline = (open.baseline + c.originY) / 2
    }
  }
  flush()
  return lines
}
