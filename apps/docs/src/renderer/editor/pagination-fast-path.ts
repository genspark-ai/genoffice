import type { Transaction } from '@tiptap/pm/state'
import { AddMarkStep, RemoveMarkStep, ReplaceStep } from '@tiptap/pm/transform'
import type { BlockBox, PageSlice } from '../pagination-types'

/**
 * Word-style fast repagination check. A keystroke inside one paragraph that
 * leaves the paragraph's height unchanged cannot move any page boundary, so
 * the whole-document measure/slice pass is skipped and the previous slicing
 * stays valid (decorations already mapped through the transaction).
 */

/** top-level block types whose height alone decides their page footprint */
const FAST_BLOCK_TYPES = new Set(['docParagraph', 'docHeading', 'docListItem'])

/** DOM content that feeds pagination beyond the block's own height */
const GEOMETRY_SELECTOR =
  'img, .doc-textbox, .doc-img-wrap, .doc-note-ref, .doc-page-br, .doc-field-pagebreak, .doc-col-br, .page-gap-inline, tr'

/**
 * Top-level index of the sole paragraph-like block whose inline text a
 * transaction edited, or null when the change is anything else (block
 * insert/remove, attribute change, non-text inline nodes, several blocks).
 */
export function singleInlineEditIndex(tr: Transaction): number | null {
  if (!tr.docChanged || tr.steps.length === 0) return null
  if (tr.before.childCount !== tr.doc.childCount) return null
  let index = -1
  for (let i = 0; i < tr.steps.length; i++) {
    const step = tr.steps[i]
    const doc = tr.docs[i]
    let from: number
    let to: number
    if (step instanceof ReplaceStep) {
      if (step.slice.openStart !== 0 || step.slice.openEnd !== 0) return null
      if (!textOnly(step.slice.content)) return null
      from = step.from
      to = step.to
    } else if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
      from = step.from
      to = step.to
    } else {
      return null
    }
    if (from > to || to > doc.content.size) return null
    const $from = doc.resolve(from)
    const $to = doc.resolve(to)
    if ($from.depth < 1 || $to.depth < 1) return null
    const top = $from.index(0)
    if ($to.index(0) !== top) return null
    if (step instanceof ReplaceStep && !textOnly(doc.slice(from, to).content)) return null
    if (index !== -1 && index !== top) return null
    index = top
  }
  if (index < 0) return null
  return FAST_BLOCK_TYPES.has(tr.doc.child(index).type.name) ? index : null
}

function textOnly(content: { forEach(f: (node: { isText: boolean }) => void): void }): boolean {
  let ok = true
  content.forEach((node) => {
    if (!node.isText) ok = false
  })
  return ok
}

/** true when a page boundary falls strictly inside the block */
export function crossesPage(block: BlockBox, slices: PageSlice[]): boolean {
  const top = block.top + 0.5
  const bottom = block.top + block.height - 0.5
  if (bottom <= top) return false
  for (const s of slices) if (s.start > top && s.start < bottom) return true
  return false
}

export interface FastPassBlock {
  el: HTMLElement
  prev: BlockBox
}

/**
 * Whether the edited blocks keep the previous slicing: each must still be a
 * single-page block with no geometry-bearing content and the same measured
 * height (0.5 px tolerance, same as the slicing engine's seam tolerance).
 */
export function blocksKeepSlicing(
  edited: FastPassBlock[],
  slices: PageSlice[],
  zoomFactor: number,
): boolean {
  if (edited.length === 0 || slices.length === 0) return false
  for (const { el, prev } of edited) {
    if (prev.domHeight === undefined || crossesPage(prev, slices)) return false
    if (el.querySelector(GEOMETRY_SELECTOR)) return false
    // an empty paragraph mark flows but takes no column-balance quota: its
    // first character (or the deletion of its last) changes the slicing at
    // the same height (same test as measureBlocks; images are rejected above)
    if (!(el.textContent ?? '').trim() !== (prev.emptyPara ?? false)) return false
    const height = el.getBoundingClientRect().height / zoomFactor
    if (Math.abs(height - prev.domHeight) > 0.5) return false
  }
  return true
}

/**
 * Index (in tr.doc) of the first top-level block a transaction changed: the
 * pages above it paginate as before. null when a step carries no position
 * (document attributes), which means the whole document.
 */
export function firstChangedTopLevelIndex(tr: Transaction): number | null {
  const maps = tr.mapping.maps
  let min = Infinity
  let unknown = false
  tr.steps.forEach((step, i) => {
    const carry = (pos: number) => {
      let p = pos
      for (let j = i + 1; j < maps.length; j++) p = maps[j].map(p, -1)
      return p
    }
    let moved = false
    maps[i].forEach((_oldStart, _oldEnd, newStart) => {
      moved = true
      min = Math.min(min, carry(newStart))
    })
    if (moved) return
    const s = step as { from?: unknown; pos?: unknown }
    if (typeof s.from === 'number') min = Math.min(min, carry(s.from))
    else if (typeof s.pos === 'number') min = Math.min(min, carry(s.pos))
    else unknown = true
  })
  if (unknown || min === Infinity) return null
  return tr.doc.resolve(Math.max(0, Math.min(min, tr.doc.content.size))).index(0)
}
