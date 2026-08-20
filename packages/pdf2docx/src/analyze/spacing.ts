/**
 * before_space vertical positioning chain (pdf2docx rule 4, second half):
 * absolute block positions become flowable by turning inter-block whitespace
 * into the LOWER block's spacingBefore. Rebuilt paragraphs occupy exactly
 * their measured ink extent (the rebuild layer writes explicit exact line
 * heights, P5), so the FULL ink-to-ink gap is emitted; negative gaps
 * (overlapping blocks) clamp to 0 with a warning.
 *
 * Page-top leading is NOT set here — it depends on the final page margins,
 * which only the rebuild layer knows (it adds `contentTop − firstBlock.top`).
 */
import type { PageBlock, PageSection } from '../ir'

/** spacing under this many points is noise — leave the field unset */
const EMIT_MIN_PT = 2
/** gaps more negative than this raise an overlap warning */
const NEG_TOL_PT = 1

const isFloat = (b: PageBlock): boolean => b.kind === 'image' && b.float !== undefined

/** hidden micro text kept for fidelity (P11 C) must not cost flow space:
 * chaining a real gap onto an invisible block pushes visible content off
 * the page (its own ~1pt lines are the only height it may consume) */
const MICRO_TEXT_MAX_PT = 2.5
export const isMicroTextBlock = (b: PageBlock): boolean =>
  b.kind === 'text' &&
  b.lines.length > 0 &&
  b.lines.every((l) => l.spans.every((s) => s.fontSize < MICRO_TEXT_MAX_PT))

/**
 * Fill spacingBeforePt across a page's sections. Chain rules:
 * - inside a column, each block's gap to the block above it;
 * - the reading-order FIRST block of a section chains from the previous
 *   section's bottom (columns 2+ start mid-flow in docx — no spacing there);
 * - floats are outside the flow and neither give nor receive spacing.
 * Returns warnings (negative gaps clamped).
 */
export function applySpacingChain(sections: readonly PageSection[]): string[] {
  const warnings: string[] = []
  let prevSectionBottom: number | null = null

  for (const section of sections) {
    for (const [ci, column] of section.columns.entries()) {
      const flow = column.blocks.filter((b) => !isFloat(b) && !isMicroTextBlock(b))
      for (const [i, block] of flow.entries()) {
        let gap: number | null = null
        if (i > 0) {
          gap = flow[i - 1]!.box.y0 - block.box.y1
        } else if (ci === 0 && prevSectionBottom !== null) {
          gap = prevSectionBottom - block.box.y1
        }
        if (gap === null) continue
        if (gap < -NEG_TOL_PT) {
          warnings.push(`overlapping blocks: negative gap ${gap.toFixed(1)}pt clamped to 0`)
          gap = 0
        }
        if (gap >= EMIT_MIN_PT) block.spacingBeforePt = gap
      }
    }
    prevSectionBottom = section.box.y0
  }
  return warnings
}
