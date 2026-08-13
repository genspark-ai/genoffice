/**
 * Vertical layout for the WPS-style comments margin: each card wants to sit level
 * with its pin, cards must not overlap, and the anchored card (active thread or the
 * draft being typed) keeps its exact pin-aligned position while the others yield.
 */

/** Offset from a pin's center down to the card row its leader line attaches to */
export const CARD_PIN_ALIGN = 14

export interface MarginLayoutEntry {
  key: string
  /** Pin center y, relative to the margin column top */
  pinY: number
}

/**
 * Compute card tops. Entries are laid out in pin order (top to bottom). Without an
 * anchor a single downward pass pushes overlapping cards down. With an anchor, the
 * anchored card is fixed at its pin-aligned top, cards above it are pushed up and
 * cards below pushed down as needed.
 */
export function layoutMarginCards(
  entries: MarginLayoutEntry[],
  height: (key: string) => number,
  anchorKey: string | null,
  gap = 10,
  minTop = 4,
): Map<string, number> {
  const sorted = [...entries].sort((a, b) => a.pinY - b.pinY)
  const desired = (e: MarginLayoutEntry) => Math.max(minTop, e.pinY - CARD_PIN_ALIGN)
  const tops = new Array<number>(sorted.length)

  const anchorIdx = anchorKey === null ? -1 : sorted.findIndex((e) => e.key === anchorKey)
  if (anchorIdx < 0) {
    for (let i = 0; i < sorted.length; i++) {
      const min = i === 0 ? minTop : tops[i - 1]! + height(sorted[i - 1]!.key) + gap
      tops[i] = Math.max(desired(sorted[i]!), min)
    }
  } else {
    tops[anchorIdx] = desired(sorted[anchorIdx]!)
    // Above the anchor: pull cards up when they would collide (may go past minTop —
    // avoiding overlap wins over staying inside the top edge)
    for (let i = anchorIdx - 1; i >= 0; i--) {
      const max = tops[i + 1]! - gap - height(sorted[i]!.key)
      tops[i] = Math.min(desired(sorted[i]!), max)
    }
    for (let i = anchorIdx + 1; i < sorted.length; i++) {
      const min = tops[i - 1]! + height(sorted[i - 1]!.key) + gap
      tops[i] = Math.max(desired(sorted[i]!), min)
    }
  }

  const out = new Map<string, number>()
  for (let i = 0; i < sorted.length; i++) out.set(sorted[i]!.key, tops[i]!)
  return out
}
