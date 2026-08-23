/**
 * 2.3 Text metrics (fidelity linchpin) — font metrics abstraction layer.
 *
 * The key to fidelity is owning the font metrics: don't rely on browser Canvas
 * measureText (untestable, inconsistent across environments); instead parse real
 * font files with opentype.js to get exact advance width / ascent / descent, so
 * any environment (including node unit tests) yields deterministic wrapping and
 * line-height results.
 *
 * Design: FontMetricsProvider interface + two implementations:
 *   1. OpentypeMetrics — real fonts (opentype.Font), pixel-accurate.
 *   2. HeuristicMetrics — deterministic fallback without font files (advance
 *      estimated by character class), keeping wrapping logic unit-testable and
 *      cross-environment consistent; once the frontend loads real fonts it
 *      switches to exact metrics automatically.
 */

export interface RunStyle {
  fontFamily: string
  fontSizePx: number
  bold: boolean
  italic: boolean
}

export interface FontMetrics {
  /** Ascent at the font size (baseline to top, px) */
  ascent: number
  /** Descent at the font size (baseline to bottom, px, positive) */
  descent: number
  /** Suggested line height (px) — the line box, without external leading */
  lineHeight: number
  /** External leading (hhea lineGap, px): advances the next line but sits outside the
   *  line box (CoreText/PowerPoint single-spacing = lineHeight + externalLeading) */
  externalLeading?: number
}

export interface FontMetricsProvider {
  /** Line metrics for this font + size. */
  metrics(style: RunStyle): FontMetrics
  /** Advance width of a piece of text under this style (px). */
  measure(text: string, style: RunStyle): number
  /**
   * CSS font family actually used for drawing. When a missing font is substituted
   * by an installed one, return the substitute name so the renderer draws with the
   * same font file the metrics used (otherwise baked run x positions drift from the
   * real widths of the browser's own fallback font → overlap/squeezing).
   * Unimplemented = use the original name. `text` lets implementations adjust per
   * script (complex-script runs measure/draw with the same shaping font).
   */
  displayFamily?(style: RunStyle, text?: string): string
}

// ── Grapheme clusters ───────────────────────────────────────────────

const SEGMENTER: Intl.Segmenter | null =
  typeof Intl !== 'undefined' && 'Segmenter' in Intl
    ? new Intl.Segmenter(undefined, { granularity: 'grapheme' })
    : null

/**
 * Split by grapheme cluster (combining marks / ZWJ emoji / flags / skin-tone
 * sequences stay intact). Both measurement and line breaking must treat clusters
 * as the smallest unit — splitting by code point would push Thai combining vowels
 * to a line start and measure an emoji sequence as N character widths.
 */
export function graphemes(text: string): string[] {
  if (!SEGMENTER) return [...text]
  const out: string[] = []
  for (const s of SEGMENTER.segment(text)) out.push(s.segment)
  return out
}

// ── Heuristic fallback (deterministic, unit-testable) ───────────────

/** Unicode East Asian Width = Wide/Fullwidth ranges (covers all fullwidth scripts at once; don't patch per language). */
const WIDE_RANGES: ReadonlyArray<readonly [number, number]> = [
  [0x1100, 0x115f], // Hangul Jamo (initial consonants)
  [0x2e80, 0x303e], // CJK radicals / Kangxi radicals / CJK punctuation
  [0x3041, 0x33ff], // Kana / Bopomofo / compatibility Hangul Jamo / CJK compatibility
  [0x3400, 0x4dbf], // CJK Extension A
  [0x4e00, 0x9fff], // CJK basic
  [0xa000, 0xa4cf], // Yi
  [0xa960, 0xa97f], // Hangul Jamo Extended-A
  [0xac00, 0xd7a3], // Hangul syllables
  [0xf900, 0xfaff], // compatibility ideographs
  [0xfe30, 0xfe4f], // CJK compatibility forms
  [0xff00, 0xff60], // fullwidth ASCII
  [0xffe0, 0xffe6], // fullwidth symbols
  [0x1b000, 0x1b2ff], // Kana extensions
  [0x20000, 0x3fffd], // CJK Extension B+
]

/** Whether a code point is EAW fullwidth/wide (word breaking may break per cluster on these). */
export function isWideChar(code: number): boolean {
  for (const [lo, hi] of WIDE_RANGES) {
    if (code >= lo && code <= hi) return true
  }
  return false
}

/**
 * Estimate advance by character class (ratio relative to font size).
 * EAW fullwidth ≈ 1.0em, general Latin ≈ 0.5em, narrow chars smaller. Bold adds a bit.
 * These factors are only for the deterministic no-real-font fallback and don't aim
 * for pixel accuracy; real typesetting is handled by OpentypeMetrics.
 */
function charAdvanceEm(code: number): number {
  if (isWideChar(code)) return 1.0
  // emoji ranges (rough)
  if (code >= 0x1f000 || (code >= 0x2600 && code <= 0x27bf)) return 1.0
  // Geometric shapes (◆ ○ ▪ ► …) and the CJK reference mark ※: EAW "Ambiguous" —
  // this fallback runs when the resolved font lacks the glyph, and the browser then
  // substitutes a CJK font where these draw full-width. Over-estimating only widens
  // a gap; under-estimating makes bullet glyphs overlap the text they precede.
  if ((code >= 0x25a0 && code <= 0x25ff) || code === 0x203b) return 1.0
  // narrow characters
  if ("iIlj.,:;'!|".includes(String.fromCharCode(code))) return 0.28
  if (' ftr'.includes(String.fromCharCode(code))) return 0.32
  if ('mwMW'.includes(String.fromCharCode(code))) return 0.82
  // general Latin / digits
  return 0.52
}

/** Zero-width code points: non-spacing combining marks (Thai vowels/tones, diacritics) and format chars (ZWJ etc.). */
const ZERO_WIDTH_RE = /[\p{Mn}\p{Me}\p{Cf}]/u
/** Emoji composition parts: ZWJ (200d) / VS16 (fe0f) / regional indicators (flags) / skin-tone modifiers. */
// eslint-disable-next-line no-misleading-character-class -- individual composition code points are matched on purpose
const EMOJI_JOIN_RE = /[\u200d\ufe0f\u{1f1e6}-\u{1f1ff}\u{1f3fb}-\u{1f3ff}]/u

/**
 * Estimated width of one grapheme cluster: an emoji composition cluster draws as
 * one glyph at 1em; otherwise sum per code point, counting zero-width marks as 0
 * (spacing vowel marks Mc, e.g. Devanagari, have real width and use the default factor).
 */
function clusterAdvanceEm(g: string): number {
  if (g.length > 1 && EMOJI_JOIN_RE.test(g)) return 1.0
  let em = 0
  for (const ch of g) {
    em += ZERO_WIDTH_RE.test(ch) ? 0 : charAdvanceEm(ch.codePointAt(0) ?? 0)
  }
  return em
}

export class HeuristicMetrics implements FontMetricsProvider {
  metrics(style: RunStyle): FontMetrics {
    const s = style.fontSizePx
    // Typical sans-serif ratios: ascent≈0.8em, descent≈0.2em, lineHeight≈1.2em
    return {
      ascent: s * 0.8,
      descent: s * 0.2,
      lineHeight: s * 1.2,
    }
  }

  measure(text: string, style: RunStyle): number {
    let em = 0
    for (const g of graphemes(text)) {
      em += clusterAdvanceEm(g)
    }
    const boldFactor = style.bold ? 1.04 : 1
    return em * style.fontSizePx * boldFactor
  }
}

// ── Exact metrics via opentype.js ───────────────────────────────────

/** Minimal interface of opentype.Font (avoids a hard type dependency; the frontend injects the real instance). */
export interface OpentypeFontLike {
  unitsPerEm: number
  ascender: number
  descender: number
  /** hhea lineGap (external leading, font units); part of the single-spacing line height */
  lineGap?: number
  getAdvanceWidth(text: string, fontSize: number): number
  /** Optional: char → glyph index (0 = missing glyph). Used for the missing-glyph heuristic fallback. */
  charToGlyphIndex?(char: string): number
}

/**
 * Exact metrics using fonts already loaded via opentype.js.
 * fontResolver: returns the font instance for family+bold+italic (undefined if not found → caller falls back).
 */
export class OpentypeMetrics implements FontMetricsProvider {
  constructor(
    private fontResolver: (style: RunStyle) => OpentypeFontLike | undefined,
    private fallback: FontMetricsProvider = new HeuristicMetrics(),
  ) {}

  metrics(style: RunStyle): FontMetrics {
    const font = this.fontResolver(style)
    if (!font) return this.fallback.metrics(style)
    const scale = style.fontSizePx / font.unitsPerEm
    const ascent = font.ascender * scale
    const descent = Math.abs(font.descender) * scale
    // Single spacing includes the hhea lineGap as EXTERNAL leading: CoreText — and with
    // it PowerPoint for Mac — paces Times New Roman/Arial at (asc+desc+gap) = 1.15em,
    // not the bare asc+desc 1.107/1.117em (dropping the gap made multi-line text creep
    // upward). The gap advances lines but stays outside the line box.
    const lineGap = (font.lineGap ?? 0) * scale
    return {
      ascent,
      descent,
      lineHeight: ascent + descent,
      ...(lineGap > 0 ? { externalLeading: lineGap } : {}),
    }
  }

  measure(text: string, style: RunStyle): number {
    const font = this.fontResolver(style)
    if (!font) return this.fallback.measure(text, style)
    try {
      // Missing-glyph guard: if any character has no glyph (e.g. CJK landing on a Latin
      // font) → fall back to heuristics for the whole run, so advance=0 doesn't make
      // line widths badly undersized.
      if (font.charToGlyphIndex) {
        for (const ch of text) {
          if (font.charToGlyphIndex(ch) === 0) return this.fallback.measure(text, style)
        }
      }
      return font.getAdvanceWidth(text, style.fontSizePx)
    } catch {
      return this.fallback.measure(text, style)
    }
  }
}
