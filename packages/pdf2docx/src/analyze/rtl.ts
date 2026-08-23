/**
 * RTL support (plan §5.3): PDF chars arrive in VISUAL order (glyph placement);
 * docx stores LOGICAL order. Per line: sort by x, infer word gaps while the
 * geometry is still meaningful, then run the UBA (bidi-js — never a hand-rolled
 * UAX#9) to recover logical order. Applying the UBA's visual reordering to a
 * visual-order sequence is the standard involution trick: it restores logical
 * order for the common cases (pure RTL, embedded digit runs, RTL/LTR mixes).
 */
import bidiFactory, { type Bidi } from 'bidi-js'
import type { Dir, Line, PdfChar } from '../ir'
import { rectCenterX } from '../geometry'
import { isRtlScript } from '../script'
import { isSpaceCode } from './lines'
import { medianCharGap, spaceGapThreshold } from './words'

let bidiSingleton: Bidi | null = null
const getBidi = (): Bidi => (bidiSingleton ??= bidiFactory())

// Arabic presentation forms (ONLY these ranges get NFKC — full-text NFKC would
// mangle compatibility chars of other scripts)
const isArabicPresentationForm = (code: number): boolean =>
  (code >= 0xfb50 && code <= 0xfdff) || (code >= 0xfe70 && code <= 0xfeff)

/**
 * Fold Arabic presentation-form code points (ligatures, positional variants)
 * back to base letters via NFKC, applied per matching char only.
 */
export function normalizeArabicForms(chars: readonly PdfChar[]): PdfChar[] {
  return chars.map((c) => {
    if (!isArabicPresentationForm(c.code)) return c
    let text = c.text.normalize('NFKC')
    // isolated harakat normalize to "space + combining mark"; drop the space
    if (text.length > 1 && text.startsWith(' ')) text = text.slice(1)
    return text === c.text ? c : { ...c, text }
  })
}

export const lineHasRtl = (chars: readonly PdfChar[]): boolean =>
  chars.some((c) => isRtlScript(c.script))

/**
 * Broken-ToUnicode ligature repair (P29 D): some Arabic fonts map the lam-alef
 * ligature glyphs to junk ASCII letters ('T', 'J', …), which then ride into
 * the docx as literal Latin letters mid-word. A single ASCII letter that is
 * x-adjacent to Arabic glyphs FROM THE SAME FONT (Arabic on one side, Arabic
 * or a word boundary on the other) can only be such a glyph — real embedded
 * Latin comes as whole words in a Latin font. The bare lam-alef stands in
 * (the hamza variant is unrecoverable, and a missing hamza reads fine).
 */
export function repairArabicJunkLigatures(visual: PdfChar[], lineGap: number): void {
  for (let i = 0; i < visual.length; i++) {
    const c = visual[i]!
    if (c.text.length !== 1 || !/[A-Za-z]/.test(c.text)) continue
    const gapTo = (a: PdfChar, b: PdfChar): number => b.looseBox.x0 - a.looseBox.x1
    const threshold = spaceGapThreshold(Math.max(c.fontSize, 1), lineGap)
    const arabicNeighbor = (n: PdfChar | undefined, gap: number): boolean =>
      n !== undefined && n.script === 'arabic' && n.fontFamily === c.fontFamily && gap < threshold
    const left = visual[i - 1]
    const right = visual[i + 1]
    const leftArabic = arabicNeighbor(left, left ? gapTo(left, c) : Infinity)
    const rightArabic = arabicNeighbor(right, right ? gapTo(c, right) : Infinity)
    const leftBoundary = left === undefined || isSpaceCode(left.code) || gapTo(left, c) >= threshold
    const rightBoundary =
      right === undefined || isSpaceCode(right.code) || gapTo(c, right) >= threshold
    if (
      (leftArabic || rightArabic) &&
      (leftArabic || leftBoundary) &&
      (rightArabic || rightBoundary)
    ) {
      visual[i] = { ...c, text: 'لا', code: 0x0644, script: 'arabic' }
    }
  }
}

/** synthetic inter-word space (materialized while geometry is still visual) */
function syntheticSpace(at: number, template: PdfChar): PdfChar {
  return {
    ...template,
    code: 0x20,
    text: ' ',
    box: { x0: at, x1: at, y0: template.box.y0, y1: template.box.y1 },
    looseBox: { x0: at, x1: at, y0: template.looseBox.y0, y1: template.looseBox.y1 },
    originX: at,
    isGenerated: true,
    isHyphen: false,
    script: 'common',
  }
}

/**
 * One RTL-bearing line: visual (any input order) → logical char order.
 * Word gaps are inferred in visual order and materialized as real space chars,
 * so the word layer must NOT re-infer them geometrically (gaps between logical
 * neighbours are meaningless once the sequence is reordered).
 */
export function reorderVisualToLogical(chars: readonly PdfChar[]): PdfChar[] {
  // 1. visual order = left → right
  const visual = [...chars].sort((a, b) => rectCenterX(a.looseBox) - rectCenterX(b.looseBox))

  // 2. materialize inferred word gaps while x-adjacency still means adjacency
  const lineGap = medianCharGap(visual)
  repairArabicJunkLigatures(visual, lineGap)
  const withSpaces: PdfChar[] = []
  for (const c of visual) {
    const prev = withSpaces[withSpaces.length - 1]
    if (
      prev &&
      !isSpaceCode(prev.code) &&
      !isSpaceCode(c.code) &&
      c.looseBox.x0 - prev.looseBox.x1 >
        spaceGapThreshold(Math.max(Math.min(prev.fontSize, c.fontSize), 1), lineGap)
    ) {
      withSpaces.push(syntheticSpace(prev.looseBox.x1, prev))
    }
    withSpaces.push(c)
  }

  // 3. UBA over the visual string; map code units back to char clusters
  const unitToChar: number[] = []
  let text = ''
  for (const [i, c] of withSpaces.entries()) {
    const t = c.text || ' '
    for (let u = 0; u < t.length; u++) unitToChar.push(i)
    text += t
  }
  const bidi = getBidi()
  const levels = bidi.getEmbeddingLevels(text)
  const indices = bidi.getReorderedIndices(text, levels)

  const out: PdfChar[] = []
  const seen = new Set<number>()
  for (const unit of indices) {
    const charIdx = unitToChar[unit]!
    if (seen.has(charIdx)) continue
    seen.add(charIdx)
    let c = withSpaces[charIdx]!
    // visually-mirrored pairs (brackets…) must flip back for logical storage.
    // (getMirroredCharacter directly — getMirroredCharactersMap has an upstream
    // bug indexing the levels result object instead of its .levels array)
    if (c.text.length === 1 && (levels.levels[unit]! & 1) === 1) {
      const mirror = bidi.getMirroredCharacter(c.text)
      if (mirror) c = { ...c, text: mirror }
    }
    out.push(c)
  }
  return out
}

/** Paragraph base direction: first strong-directional character wins. */
export function firstStrongDir(lines: readonly Line[]): Dir {
  for (const line of lines) {
    for (const span of line.spans) {
      if (isRtlScript(span.script)) return 'rtl'
      if (span.script !== 'common') return 'ltr'
    }
  }
  return 'ltr'
}
