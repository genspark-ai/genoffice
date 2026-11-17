import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import { ReplaceAroundStep, ReplaceStep } from '@tiptap/pm/transform'

/**
 * Ranges a transaction touched, as [from, to] in its final document; null
 * when a step is of a kind this cannot place, so the caller must rescan.
 */
export function touchedRanges(tr: Transaction): Array<[number, number]> | null {
  const maps = tr.mapping.maps
  const ranges: Array<[number, number]> = []
  let unknown = false
  tr.steps.forEach((step, i) => {
    // positions after step i, carried through the later steps into the final doc
    const push = (start: number, end: number) => {
      let from = start
      let to = end
      for (let j = i + 1; j < maps.length; j++) {
        from = maps[j].map(from, -1)
        to = maps[j].map(to, 1)
      }
      ranges.push([from, to])
    }
    let moved = false
    maps[i].forEach((_oldStart, _oldEnd, newStart, newEnd) => {
      moved = true
      push(newStart, newEnd)
    })
    if (moved) return
    // mark and attribute steps leave positions alone but change the nodes
    const s = step as { from?: unknown; to?: unknown; pos?: unknown }
    if (typeof s.from === 'number' && typeof s.to === 'number') push(s.from, s.to)
    else if (typeof s.pos === 'number') push(s.pos, s.pos + 1)
    else unknown = true
  })
  return unknown ? null : ranges
}

/**
 * Top-level blocks a transaction touched, as offsets in its final document;
 * null when a step replaces so much that recomputing everything is cheaper.
 */
export function touchedTopLevelBlocks(tr: Transaction): Set<number> | null {
  const ranges = touchedRanges(tr)
  if (!ranges) return null
  const { doc } = tr
  const touched = new Set<number>()
  doc.forEach((node, offset) => {
    const end = offset + node.nodeSize
    if (ranges.some(([from, to]) => offset < to && end > from)) touched.add(offset)
  })
  return touched.size * 2 > doc.childCount ? null : touched
}

/** whether any range the transaction touched intersects a node matching the
    predicate in its final document. Unknown step kinds answer "yes" — the
    safe direction, since callers use this to skip a whole-document walk. */
export function touchesNode(
  tr: Transaction,
  predicate: (node: ProseMirrorNode) => boolean,
): boolean {
  const ranges = touchedRanges(tr)
  if (!ranges) return true
  for (const range of ranges) {
    const from = Math.min(range[0], range[1])
    const to = Math.max(range[0], range[1])
    let hit = false
    tr.doc.nodesBetween(from, to, (node) => {
      if (predicate(node)) {
        hit = true
        return false
      }
      return true
    })
    if (hit) return true
  }
  return false
}

/** the text a transaction inserted or replaced, in its final document; null
    when a step is of a kind this cannot place */
export function touchedText(tr: Transaction): string | null {
  const ranges = touchedRanges(tr)
  if (!ranges) return null
  let out = ''
  for (const range of ranges) {
    const from = Math.min(range[0], range[1])
    const to = Math.max(range[0], range[1])
    if (to > from) out += tr.doc.textBetween(from, to, ' ')
  }
  return out
}

/** whether the transaction removed any content: deletions cannot introduce a
    new node, but they can remove the last instance of one that was there.
    A cross-block deletion lands as ReplaceAroundStep (its outer range minus
    the retained gap always removes the boundaries). */
export function deletedAnything(tr: Transaction): boolean {
  return tr.steps.some(
    (s) =>
      (s instanceof ReplaceStep && s.to > s.from) ||
      (s instanceof ReplaceAroundStep && s.from < s.to),
  )
}

/** a single insertion at the document end: every existing position survives unchanged */
export function appendsAtEnd(tr: Transaction): boolean {
  if (tr.steps.length !== 1) return false
  const step = tr.steps[0]
  if (!(step instanceof ReplaceStep)) return false
  return step.from === step.to && step.from === tr.before.content.size
}
