/**
 * Pure math behind PowerPoint-style multi-selection in the thumbnail rail and
 * sorter: click / Shift-click / Ctrl-click selection, and the op plans that
 * turn a multi-slide duplicate or drag-reorder into one transaction.
 */

export interface SlideSelection {
  /** Ascending, never empty, always contains `current` */
  selected: number[]
  /** Anchor slide: shown on the canvas, start of Shift ranges */
  current: number
}

export function rangeSelection(anchor: number, target: number): number[] {
  const lo = Math.min(anchor, target)
  const hi = Math.max(anchor, target)
  return Array.from({ length: hi - lo + 1 }, (_, k) => lo + k)
}

function sortUnique(indexes: number[]): number[] {
  return [...new Set(indexes)].sort((a, b) => a - b)
}

/** Ctrl/Cmd-click. Dropping the anchor hands it to the nearest slide still selected; the last slide cannot be dropped. */
export function toggleSelection(sel: SlideSelection, index: number): SlideSelection {
  if (!sel.selected.includes(index)) {
    return { selected: sortUnique([...sel.selected, index]), current: sel.current }
  }
  if (sel.selected.length === 1) return sel
  const selected = sel.selected.filter((i) => i !== index)
  if (index !== sel.current) return { selected, current: sel.current }
  const current = selected.reduce((best, i) =>
    Math.abs(i - index) < Math.abs(best - index) ? i : best,
  )
  return { selected, current }
}

export function clickSelection(
  sel: SlideSelection,
  index: number,
  mods: { shift?: boolean; toggle?: boolean },
): SlideSelection {
  if (mods.shift) return { selected: rangeSelection(sel.current, index), current: sel.current }
  if (mods.toggle) return toggleSelection(sel, index)
  return { selected: [index], current: index }
}

/** Selection to render for the given anchor/deck size; returns `selected` itself while it is still valid. */
export function normalizeSelection(selected: number[], current: number, count: number): number[] {
  const valid =
    selected.length > 0 && selected.includes(current) && selected.every((i) => i >= 0 && i < count)
  return valid ? selected : [current]
}

/** Keep the current slide across a history restore even when its position changed. */
export function currentAfterHistory(
  slides: readonly { partPath?: string }[],
  current: number,
  partPath?: string,
): number {
  if (partPath) {
    const index = slides.findIndex((slide) => slide.partPath === partPath)
    if (index >= 0) return index
  }
  return Math.min(current, slides.length - 1)
}

/** Where the moved slides sit after landing as a block at insertAt (a gap 0..count in the pre-move order). */
export function movedBlockPositions(selected: number[], insertAt: number): number[] {
  const first = insertAt - selected.filter((i) => i < insertAt).length
  return selected.map((_, k) => first + k)
}

export type MoveStep = { from: number; to: number }

/**
 * moveSlide steps (each validated against the pre-transaction deck, applied
 * in sequence) that land `selected` as a block at gap insertAt. Empty when
 * the order would not change.
 */
export function planSlideMoves(count: number, selected: number[], insertAt: number): MoveStep[] {
  const block = sortUnique(selected)
  const picked = new Set(block)
  const rest = Array.from({ length: count }, (_, i) => i).filter((i) => !picked.has(i))
  const at = insertAt - block.filter((i) => i < insertAt).length
  const final = [...rest.slice(0, at), ...block, ...rest.slice(at)]
  const cur = Array.from({ length: count }, (_, i) => i)
  const steps: MoveStep[] = []
  final.forEach((id, t) => {
    const from = cur.indexOf(id)
    if (from === t) return
    steps.push({ from, to: t })
    cur.splice(from, 1)
    cur.splice(t, 0, id)
  })
  return steps
}

export type DuplicateStep = { duplicate: number } | MoveStep

/**
 * Copies land in order right after the last selected slide (PowerPoint). The
 * last one is duplicated first so no step targets an index that does not
 * exist in the pre-transaction deck.
 */
export function planSlideDuplicates(selected: number[]): DuplicateStep[] {
  const block = sortUnique(selected)
  const last = block[block.length - 1]
  if (last == null) return []
  const steps: DuplicateStep[] = [{ duplicate: last }]
  block.slice(0, -1).forEach((i, j) => {
    steps.push({ duplicate: i }, { from: i + 1, to: last + 1 + j })
  })
  return steps
}
