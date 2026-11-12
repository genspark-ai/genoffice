/**
 * Tab of a display-only stray line (the anchor paragraph's own text, rendered
 * as static DOM without the measuring tab-stop plugin). The emitter cannot
 * measure the text before the tab, so the rule decides from the stops alone.
 *
 * Word (probe 2026-11-07, TOC entries `title<tab>page` with a right stop 130
 * twips inside the column edge, on it, and 144 twips past it): the page number
 * ends on the stop whatever its width; a stop past the column edge is laid on
 * the edge; a title that passes the stop collapses the tab and the number
 * follows it, wrapping only when it would cross the edge.
 */
import type { TabStop } from '@genoffice/docx-engine'

export type StrayTabPlan =
  /** Chromium tab-size grid from the content edge: the first stop past the indent sizes it */
  | { kind: 'grid'; tabSizeTwips: number }
  /** one right/decimal stop: the text after the tab ends on the stop, clamped at the column edge */
  | { kind: 'right'; stopTwips: number }

/** @param tabCount tab characters on the stray line (a right box can hold one segment) */
export function planStrayTab(stops: TabStop[], leftTwips: number, tabCount: number): StrayTabPlan {
  const next = stops
    .filter((s) => s.val !== 'clear' && s.val !== 'bar' && !s.rel && s.pos > leftTwips)
    .sort((a, b) => a.pos - b.pos)[0]
  if (
    next &&
    tabCount === 1 &&
    (next.val === 'right' || next.val === 'decimal') &&
    (!next.leader || next.leader === 'none')
  )
    return { kind: 'right', stopTwips: next.pos }
  return { kind: 'grid', tabSizeTwips: (next?.pos ?? leftTwips + 720) - leftTwips }
}
