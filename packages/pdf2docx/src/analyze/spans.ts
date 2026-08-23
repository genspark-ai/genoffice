/**
 * words → spans. A span is the docx run precursor: uniform font family/size,
 * bold/italic, color AND script. Script changes split spans (mixed zh/en text
 * produces alternating CJK/Latin spans so the rebuild layer can map font
 * slots per script); 'common' chars (digits, punctuation) attach to the
 * current span's script.
 */
import type { Rect } from '../geometry'
import { approxEq, median, rectUnion } from '../geometry'
import type { Dir, PdfChar, Span } from '../ir'
import type { UnicodeScript } from '../script'
import { isRtlScript } from '../script'
import type { Word } from './words'

/** font sizes within this many points are the same style */
const FONT_SIZE_TOL = 0.1

// ── character-compression restore (P5) ──
/** glyph scales this close to 1 are noise, not w:w */
const CHAR_SCALE_TOL = 0.015
/** per-char tracking under this many points is noise, not w:spacing */
const CHAR_SPACING_MIN_PT = 0.05
/** advance deltas outside (0, 2em] are line/word jumps — not tracking samples */
const CHAR_SPACING_MAX_EMS = 2
/**
 * ink gap across a real space char at/above this many ems reads as a normal
 * word gap (P14 B). PowerPoint's PDF writer declares /Widths wider than the
 * advances it lays out and pulls every glyph pair back with TJ kerns; the
 * measured "tracking" is then a large negative constant although the text is
 * NOT visually squeezed. Restored as w:spacing it crushes the (normal-width)
 * spaces of the substituted font and the words glue together — healthy space
 * ink gaps prove the negative median is a metrics artifact, so drop it.
 * Genuinely squeezed text squeezes its spaces too and fails this bar.
 */
const SPACE_INK_HEALTHY_EMS = 0.18
/**
 * Italic/swash fonts overhang their ink boxes: adjacent glyphs' boxes overlap
 * (negative intra-word ink gaps) and the same overhang shaves the measured
 * space ink gap below the healthy bar although the words read normally
 * (P15 A: HTML-export decks with inflated /Widths + an italic serif face).
 * Adding the median intra-word overlap back recovers the true visual space
 * width. Only tracking at least this negative (in ems) trusts the corrected
 * gap — mild genuine compression keeps its spacing.
 */
const ARTIFACT_TRACKING_MIN_EMS = 0.15

// \bdemi\b: 'ITC Franklin Gothic Std Demi' is a bold-weight face; the word
// boundary keeps 'Academi…'-style substrings out
const BOLD_NAME = /bold|black|heavy|semibold|demibold|\bdemi\b/i

export const isBoldChar = (c: PdfChar): boolean =>
  c.fontWeight >= 600 || BOLD_NAME.test(c.fontFamily)

function sameStyle(a: PdfChar, b: PdfChar): boolean {
  return (
    a.fontFamily === b.fontFamily &&
    approxEq(a.fontSize, b.fontSize, FONT_SIZE_TOL) &&
    isBoldChar(a) === isBoldChar(b) &&
    a.italic === b.italic &&
    a.color === b.color &&
    a.highlight === b.highlight &&
    a.underline === b.underline &&
    a.strike === b.strike &&
    a.invisible === b.invisible
  )
}

const dirOf = (script: UnicodeScript): Dir => (isRtlScript(script) ? 'rtl' : 'ltr')

interface OpenSpan {
  anchor: PdfChar
  lastChar: PdfChar
  text: string
  box: Rect
  script: UnicodeScript
  /** glyph scales of the span's chars (median → Span.charScale) */
  scales: number[]
  /** measured advance − font advance per adjacent pair (median → Span.charSpacingPt) */
  trackings: number[]
  /** ink gap across each real/inferred space (P14 B: metrics-artifact evidence) */
  spaceInkGaps: number[]
  /** ink gap of each tracked intra-word pair (P15 A: italic-overhang correction) */
  intraInkGaps: number[]
  /** measured advance of each lone REAL space char (P15 A: rendered space width) */
  spaceAdvances: number[]
}

/**
 * Build spans from one line's words. Inter-word spaces become literal ' '
 * text attached to the span that precedes them.
 */
export function buildSpans(words: readonly Word[]): Span[] {
  const spans: Span[] = []
  let open: OpenSpan | null = null

  const flush = (): void => {
    if (!open) return
    const span: Span = {
      text: open.text,
      box: open.box,
      fontSize: open.anchor.fontSize,
      fontFamily: open.anchor.fontFamily,
      bold: isBoldChar(open.anchor),
      italic: open.anchor.italic,
      color: open.anchor.color,
      dir: dirOf(open.script),
      script: open.script,
    }
    if (open.anchor.highlight) span.highlight = open.anchor.highlight
    if (open.anchor.underline) span.underline = true
    if (open.anchor.strike) span.strike = true
    if (open.anchor.invisible) span.invisible = true
    // squeezed text (AI docs: w:w scale + negative w:spacing): restore both,
    // or every rebuilt line wraps earlier than the original and pages overflow
    const scale = median(open.scales)
    if (scale > 0 && Math.abs(scale - 1) >= CHAR_SCALE_TOL) span.charScale = scale
    if (open.trackings.length >= 2) {
      const tracking = median(open.trackings)
      // negative tracking with healthy word spaces = inflated /Widths, not
      // visual squeezing (P14 B) — restoring it would glue the words
      let metricsArtifact = false
      if (tracking < 0 && (open.spaceInkGaps.length > 0 || open.spaceAdvances.length > 0)) {
        const em = Math.max(open.anchor.fontSize, 1)
        const extreme = tracking <= -(ARTIFACT_TRACKING_MIN_EMS * em)
        const spaceGap = open.spaceInkGaps.length > 0 ? median(open.spaceInkGaps) : -Infinity
        // italic-overhang correction (P15 A): the words' own ink overlap
        // measures the overhang that also shaved the space gap
        const overhang = open.intraInkGaps.length > 0 ? Math.min(0, median(open.intraInkGaps)) : 0
        // rendered width of the REAL space chars (P15 A): origin-to-origin, so
        // no ink/bearing noise — squeezed text squeezes this too, an inflated-
        // /Widths artifact leaves it at normal word-space width
        const healthyAdvance =
          open.spaceAdvances.length > 0 && median(open.spaceAdvances) >= SPACE_INK_HEALTHY_EMS * em
        metricsArtifact =
          spaceGap >= SPACE_INK_HEALTHY_EMS * em ||
          (extreme && (healthyAdvance || spaceGap - overhang >= SPACE_INK_HEALTHY_EMS * em))
      }
      if (!metricsArtifact && Math.abs(tracking) >= CHAR_SPACING_MIN_PT) {
        span.charSpacingPt = tracking
      }
    }
    spans.push(span)
    open = null
  }

  // tracking pairs span word boundaries too (each CJK char is its own word);
  // only real/inferred spaces break the chain — a space's advance is not tracking
  let prevTracked: PdfChar | null = null
  for (const word of words) {
    if (word.spaceBefore) {
      if (open) {
        open.text += ' '
        // how wide the space actually renders (ink to ink) — the P14 B
        // metrics-artifact evidence; measured before any style flush so the
        // gap stays with the span that carries the literal ' '
        const first = word.chars.find((c) => c.text.length > 0)
        if (first !== undefined) {
          open.spaceInkGaps.push(first.box.x0 - open.lastChar.box.x1)
        }
        if (word.spaceAdvancePt !== undefined && word.spaceAdvancePt > 0) {
          open.spaceAdvances.push(word.spaceAdvancePt)
        }
      }
      prevTracked = null
    }
    for (const c of word.chars) {
      // footnote anchor (P6): always its own span so the rebuild layer can
      // emit the w:footnoteReference run in place
      if (c.noteRef !== undefined) {
        flush()
        spans.push({
          text: c.text,
          box: c.box,
          fontSize: c.fontSize,
          fontFamily: c.fontFamily,
          bold: isBoldChar(c),
          italic: c.italic,
          color: c.color,
          dir: dirOf(c.script),
          script: c.script,
          noteRef: c.noteRef,
        })
        prevTracked = null
        continue
      }
      if (open) {
        const styleBreak = !sameStyle(open.lastChar, c)
        // a real (non-common) script that differs from the span's script splits;
        // a span that started on 'common' chars upgrades to the first real script
        const scriptBreak =
          c.script !== 'common' && open.script !== 'common' && c.script !== open.script
        if (styleBreak || scriptBreak) {
          flush()
          prevTracked = null
        }
      }
      if (!open) {
        open = {
          anchor: c,
          lastChar: c,
          text: '',
          box: c.box,
          script: c.script,
          scales: [],
          trackings: [],
          spaceInkGaps: [],
          intraInkGaps: [],
          spaceAdvances: [],
        }
      } else if (open.script === 'common' && c.script !== 'common') {
        open.script = c.script
      }
      open.text += c.text
      open.box = rectUnion(open.box, c.box)
      open.scales.push(c.hscale ?? 1)
      if (prevTracked) {
        // tracking sample: actual advance vs the font's own (loose box) advance
        const advance = c.originX - prevTracked.originX
        const nominal = prevTracked.looseBox.x1 - prevTracked.looseBox.x0
        if (advance > 0 && advance <= CHAR_SPACING_MAX_EMS * prevTracked.fontSize && nominal > 0) {
          open.trackings.push(advance - nominal)
          if (c.box.x1 > c.box.x0 && prevTracked.box.x1 > prevTracked.box.x0) {
            open.intraInkGaps.push(c.box.x0 - prevTracked.box.x1)
          }
        }
      }
      prevTracked = c
      open.lastChar = c
    }
  }
  flush()
  return spans
}
