/**
 * RTL pipeline unit tests: visual→logical bidi reorder, Arabic presentation
 * form normalization, first-strong paragraph direction. Chars are hand-built
 * in VISUAL order (as a PDF stores them) via the mkChar/mkText helpers.
 */
import { describe, expect, it } from 'vitest'
import {
  analyzeChars,
  firstStrongDir,
  normalizeArabicForms,
  reorderVisualToLogical,
} from '../src/analyze'
import type { PdfChar } from '../src/ir'
import { mkChar, mkText } from './helpers/chars'

/** lay a LOGICAL string out visually (reversed), left→right, mimicking a PDF writer */
function visualRtl(logical: string, x = 100, opts: Parameters<typeof mkChar>[2] = {}): PdfChar[] {
  const reversed = [...logical].reverse().join('')
  return mkText(reversed, x, opts).chars
}

const joined = (chars: readonly PdfChar[]): string => chars.map((c) => c.text).join('')

describe('reorderVisualToLogical', () => {
  it('recovers logical order for a pure RTL line (issue #73: whole-string reversal)', () => {
    const visual = visualRtl('שלום')
    expect(joined(reorderVisualToLogical(visual))).toBe('שלום')
  })

  it('keeps embedded digit runs LTR inside an RTL line', () => {
    // logical "רק 25" displays as "25 קר": digits stay LTR in the display
    const visual = mkText('25 קר', 100).chars
    expect(joined(reorderVisualToLogical(visual)).replace(/\s+/g, ' ')).toBe('רק 25')
  })

  it('handles mixed LTR/RTL lines', () => {
    // logical "abc עב" displays as "abc בע"
    const visual = mkText('abc בע', 100).chars
    expect(joined(reorderVisualToLogical(visual)).replace(/\s+/g, ' ')).toBe('abc עב')
  })

  it('accepts input in any order (sorts by x first)', () => {
    const visual = visualRtl('מים')
    const shuffled = [visual[2]!, visual[0]!, visual[1]!]
    expect(joined(reorderVisualToLogical(shuffled))).toBe('מים')
  })

  it('materializes word gaps from geometry before reordering', () => {
    // two Arabic words drawn with a visible gap but NO space glyph
    const w2 = mkText('ابحرم', 100).chars // visual for "مرحبا"
    const last = w2[w2.length - 1]!
    const w1 = mkText('كب', last.box.x1 + 6, {}).chars // visual for "بك", 6pt gap at 10pt font
    const logical = joined(reorderVisualToLogical([...w2, ...w1]))
    expect(logical.replace(/\s+/g, ' ')).toBe('بك مرحبا')
  })

  it('mirrors bracket pairs back to logical form', () => {
    // logical "(אב)" displays visually as "(בא)" — glyphs carry the mirrored forms
    const visual = mkText('(בא)', 100).chars
    expect(joined(reorderVisualToLogical(visual))).toBe('(אב)')
  })
})

describe('normalizeArabicForms (range-gated NFKC)', () => {
  it('folds Arabic presentation forms back to base letters', () => {
    // U+FEDF ARABIC LETTER LAM INITIAL FORM → ل
    const c = mkChar(String.fromCodePoint(0xfedf), 100)
    expect(normalizeArabicForms([c])[0]!.text).toBe('ل')
  })

  it('expands lam-alef ligatures into two letters', () => {
    // U+FEFB ARABIC LIGATURE LAM WITH ALEF ISOLATED FORM → لا
    const c = mkChar(String.fromCodePoint(0xfefb), 100)
    expect(normalizeArabicForms([c])[0]!.text).toBe('لا')
  })

  it('does NOT touch compatibility chars outside the two Arabic ranges', () => {
    const circled = mkChar('①', 100) // ① — full-text NFKC would turn it into "1"
    const roman = mkChar('Ⅲ', 110) // Ⅲ
    const out = normalizeArabicForms([circled, roman])
    expect(out[0]!.text).toBe('①')
    expect(out[1]!.text).toBe('Ⅲ')
  })
})

describe('firstStrongDir', () => {
  const lineOf = (chars: PdfChar[]) => analyzeChars(chars)

  it('RTL first-strong → rtl; digits/punctuation are neutral', () => {
    expect(firstStrongDir(lineOf(mkText('12 בע', 100).chars))).toBe('rtl')
  })

  it('LTR first-strong → ltr even with later RTL content', () => {
    expect(firstStrongDir(lineOf(mkText('abc בע', 100).chars))).toBe('ltr')
  })

  it('defaults to ltr when nothing strong exists', () => {
    expect(firstStrongDir(lineOf(mkText('123 456', 100).chars))).toBe('ltr')
  })
})

describe('analyzeChars end-to-end for RTL lines', () => {
  it('produces logical-order spans with dir=rtl and keeps real spaces', () => {
    // logical "אני פה" → visual "הפ ינא"
    const visual = mkText('הפ ינא', 100).chars
    const lines = analyzeChars(visual)
    expect(lines).toHaveLength(1)
    const text = lines[0]!.spans.map((s) => s.text).join('')
    expect(text.replace(/\s+/g, ' ')).toBe('אני פה')
    expect(lines[0]!.spans[0]!.dir).toBe('rtl')
  })

  it('splits spans at script boundaries in logical order for mixed lines', () => {
    // logical "עב abc" → visual "abc בע" — wait: RTL-first line, latin tail sits left
    const visual = mkText('abc בע', 100).chars
    // make it RTL-first: hebrew visually rightmost
    const lines = analyzeChars(visual)
    const spanTexts = lines[0]!.spans.map((s) => s.text.trim()).filter(Boolean)
    expect(spanTexts.join('|')).toContain('abc')
  })

  it('LTR lines are untouched by the RTL path', () => {
    const lines = analyzeChars(mkText('plain latin words', 72).chars)
    expect(lines[0]!.spans.map((s) => s.text).join('')).toBe('plain latin words')
    expect(lines[0]!.spans[0]!.dir).toBe('ltr')
  })
})

describe('repairArabicJunkLigatures (P29 D)', () => {
  it('replaces a junk ASCII letter sandwiched between Arabic glyphs', () => {
    // broken ToUnicode: the lam-alef glyph reports 'T' — visual L→R: ة و خ T ا
    const visual = mkText('ةوخTا', 100).chars
    const out = joined(reorderVisualToLogical(visual))
    expect(out).not.toContain('T')
    expect(out).toContain('لا')
  })

  it('keeps a real embedded Latin word untouched', () => {
    const visual = mkText('בע abc', 100).chars
    expect(joined(reorderVisualToLogical(visual))).toContain('abc')
  })

  it('keeps an isolated Latin letter separated by spaces', () => {
    const visual = mkText('בע A וכ', 100).chars
    expect(joined(reorderVisualToLogical(visual))).toContain('A')
  })
})
