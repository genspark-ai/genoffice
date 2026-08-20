/**
 * line chars → words. Script-aware spacing: Latin/Hangul break on real space
 * chars or on font-size-adaptive x gaps; CJK/kana/Thai characters each form
 * their own word and NEVER get machine-inserted spaces (hard rule).
 */
import type { Rect } from '../geometry'
import { median, rectUnion } from '../geometry'
import type { PdfChar } from '../ir'
import { isNoSpaceScript } from '../script'
import { isSpaceCode } from './lines'

export interface Word {
  chars: PdfChar[]
  text: string
  box: Rect
  /** separated from the previous word by a (real or inferred) space */
  spaceBefore: boolean
  /**
   * measured advance of the single REAL space char before this word (pt):
   * how wide that space actually renders (P15 A). Absent for inferred
   * spaces, generated spaces and multi-space runs.
   */
  spaceAdvancePt?: number
}

/** PDFium's generated hyphenation marker (carries no glyph; the line flag keeps it) */
const HYPHEN_MARKER = 0x02

/**
 * Adaptive inter-word gap threshold (points). Base is a fraction of the font
 * size (so it scales with the type), raised further when the line's own
 * tracking (median intra-char gap) is loose — the pdf2docx failure mode this
 * guards against is a fixed threshold misfiring on tight/loose typesetting.
 */
export function spaceGapThreshold(fontSize: number, medianCharGap: number): number {
  const base = 0.18 * fontSize
  if (medianCharGap > 0.06 * fontSize) return Math.max(base, medianCharGap + 0.1 * fontSize)
  return base
}

/** gaps above this share of the font size are word-gap candidates, not tracking */
const TRACKING_MAX_EMS = 0.15
/**
 * A PDFium-generated space at a Latin↔CJK script boundary survives the
 * no-space hard rule when it spans a real layout gap this many ems wide
 * (P20: a "5.3<tab>CJK title" heading numbering exports no tab glyph — the
 * generated space is the only witness). Justified CJK stretch never opens
 * intra-line gaps this far.
 */
const GENERATED_BOUNDARY_KEEP_EMS = 0.75

// ── letter-spaced display text (P10 C) ──
// Decorative titles position every glyph separately with a uniform tracking
// gap ("A TOAST TO THE JOURNEY AHEAD"). Those gaps must neither split the
// line into units nor become inferred spaces — the whole line stays one run
// and the spans layer restores the tracking as w:spacing.

/** intra-word adjacent-glyph gaps must be at least this many ems … */
const LETTER_SPACED_MIN_EMS = 0.5
/** … and at most this many (larger = table cells / columns, not tracking) */
const LETTER_SPACED_MAX_EMS = 2
/** minimum intra-word gaps to call the pattern (short lines prove nothing) */
const LETTER_SPACED_MIN_GAPS = 4
/** share of gaps that must sit near the median gap for "uniform" */
const LETTER_SPACED_UNIFORM_MIN = 0.75
/** a gap is "near the median" inside [0.55, 1.6] × median */
const LETTER_SPACED_BAND_LO = 0.55
const LETTER_SPACED_BAND_HI = 1.6

/**
 * True when the chars read as one letter-spaced display line: enough
 * INTRA-WORD adjacent-glyph gaps (real space chars exclude their neighbors —
 * word gaps are wider by the space's advance), all in the tracking range
 * (0.5–2 em), clustered tightly around their median. Table cell rows fail on
 * size (gaps of several ems) or on uniformity (column widths differ); normal
 * running text fails the minimum (its intra-word gaps are near zero).
 */
export function isLetterSpacedLine(chars: readonly PdfChar[]): boolean {
  const sorted = [...chars].sort((a, b) => a.looseBox.x0 - b.looseBox.x0)
  const gaps: number[] = []
  const sizes: number[] = []
  let prev: PdfChar | null = null
  let sawSpace = false
  for (const c of sorted) {
    if (isSpaceCode(c.code)) {
      sawSpace = true
      continue
    }
    if (c.code <= 0x1f) continue
    sizes.push(c.fontSize)
    if (prev && !sawSpace) gaps.push(charGap(prev, c))
    prev = c
    sawSpace = false
  }
  if (gaps.length < LETTER_SPACED_MIN_GAPS) return false
  const fontSize = Math.max(median(sizes), 1)
  const med = median(gaps)
  if (med < LETTER_SPACED_MIN_EMS * fontSize || med > LETTER_SPACED_MAX_EMS * fontSize) return false
  const near = gaps.filter(
    (g) => g >= LETTER_SPACED_BAND_LO * med && g <= LETTER_SPACED_BAND_HI * med,
  ).length
  return near / gaps.length >= LETTER_SPACED_UNIFORM_MIN
}

/**
 * median positive x gap between consecutive visible chars (the line's
 * tracking). Only clearly intra-word gaps count — a word gap sneaking into a
 * small sample would raise the threshold past itself.
 */
export function medianCharGap(chars: readonly PdfChar[]): number {
  const gaps: number[] = []
  for (let i = 1; i < chars.length; i++) {
    const prev = chars[i - 1]!
    const cur = chars[i]!
    if (isSpaceCode(prev.code) || isSpaceCode(cur.code)) continue
    const gap = charGap(prev, cur)
    const fontSize = Math.max(Math.min(prev.fontSize, cur.fontSize), 1)
    if (gap > 0 && gap <= TRACKING_MAX_EMS * fontSize) gaps.push(gap)
  }
  return median(gaps)
}

/**
 * x gap between two chars, advance-box based. The tight charBox is an INK box
 * (narrow glyphs like hangul 여 leave phantom gaps); the loose box carries the
 * advance width, so consecutive glyphs of one word touch.
 */
function charGap(prev: PdfChar, cur: PdfChar): number {
  return cur.looseBox.x0 - prev.looseBox.x1
}

function shouldInsertSpace(prev: PdfChar, cur: PdfChar, lineGap: number): boolean {
  // HARD RULE: never machine-insert spaces next to CJK/kana/Thai characters
  if (isNoSpaceScript(prev.script) || isNoSpaceScript(cur.script)) return false
  const fontSize = Math.min(prev.fontSize, cur.fontSize) || Math.max(prev.fontSize, cur.fontSize)
  return charGap(prev, cur) > spaceGapThreshold(Math.max(fontSize, 1), lineGap)
}

export interface WordOptions {
  /**
   * Infer word gaps from x geometry (default). Reordered RTL lines pass false:
   * their gaps were materialized as space chars while the sequence was still
   * visual, and x distances between logical neighbours mean nothing.
   */
  inferSpaces?: boolean
}

/** Split one line's chars into words. */
export function groupIntoWords(chars: readonly PdfChar[], options: WordOptions = {}): Word[] {
  // letter-spaced display text (P10 C): the uniform tracking gaps are NOT
  // word gaps — suppress inference so the spans layer can restore w:spacing
  const inferSpaces = (options.inferSpaces ?? true) && !isLetterSpacedLine(chars)
  const lineGap = medianCharGap(chars)
  const words: Word[] = []
  let current: PdfChar[] = []
  let pendingSpace = false
  /** every char behind the pending space was PDFium-generated (layout guess) */
  let pendingSpaceGenerated = true
  /** the run's single real space char (null once inapplicable) — its measured
   * advance is the P15 A artifact evidence */
  let pendingSpaceChar: PdfChar | null = null
  let pendingSpaceCount = 0
  let prevVisible: PdfChar | null = null

  const flush = (spaceBefore: boolean, spaceAdvancePt?: number): void => {
    if (current.length === 0) return
    const word: Word = {
      chars: current,
      text: current.map((c) => c.text).join(''),
      box: current.map((c) => c.box).reduce(rectUnion),
      spaceBefore,
    }
    if (spaceAdvancePt !== undefined) word.spaceAdvancePt = spaceAdvancePt
    words.push(word)
    current = []
  }

  let currentSpaceBefore = false
  let currentSpaceAdvance: number | undefined
  for (const c of chars) {
    if (c.code === HYPHEN_MARKER) continue
    if (isSpaceCode(c.code)) {
      if (!pendingSpace) {
        pendingSpaceGenerated = true
        pendingSpaceChar = null
        pendingSpaceCount = 0
      }
      pendingSpace = true
      pendingSpaceGenerated &&= c.isGenerated
      pendingSpaceCount++
      if (!c.isGenerated) pendingSpaceChar ??= c
      continue
    }
    // PDFium generates space chars from x-gap/regression heuristics (Kangxi
    // radicals drawn out of content order trip it). Those are machine-inserted,
    // so the CJK/kana/Thai hard rule applies to them like to our own inference;
    // real space glyphs from the content stream always stay. Exception (P20):
    // at a Latin↔CJK boundary spanning a wide real gap, the generated space
    // records a source tab/indent — dropping it fuses "5.3" into the CJK title.
    if (
      pendingSpace &&
      pendingSpaceGenerated &&
      (isNoSpaceScript(c.script) || (prevVisible !== null && isNoSpaceScript(prevVisible.script)))
    ) {
      const boundary =
        prevVisible !== null && isNoSpaceScript(prevVisible.script) !== isNoSpaceScript(c.script)
      const fontPt = Math.max(
        prevVisible !== null ? Math.min(prevVisible.fontSize, c.fontSize) : c.fontSize,
        1,
      )
      const wideGap =
        prevVisible !== null && charGap(prevVisible, c) >= GENERATED_BOUNDARY_KEEP_EMS * fontPt
      if (!(boundary && wideGap)) pendingSpace = false
    }
    const inferredSpace =
      inferSpaces &&
      prevVisible !== null &&
      current.length > 0 &&
      shouldInsertSpace(prevVisible, c, lineGap)
    const noSpaceChar = isNoSpaceScript(c.script)
    const prevNoSpace = prevVisible !== null && isNoSpaceScript(prevVisible.script)

    if (pendingSpace || inferredSpace || noSpaceChar || prevNoSpace) {
      // CJK/kana/Thai chars each stand alone; only real/inferred spaces carry a gap
      flush(currentSpaceBefore, currentSpaceAdvance)
      currentSpaceBefore = pendingSpace || inferredSpace
      // a lone real space's rendered width: origin to the next glyph's origin
      currentSpaceAdvance =
        pendingSpace && pendingSpaceCount === 1 && pendingSpaceChar !== null
          ? c.originX - pendingSpaceChar.originX
          : undefined
      pendingSpace = false
    }
    current.push(c)
    prevVisible = c
  }
  flush(currentSpaceBefore, currentSpaceAdvance)
  return words
}
