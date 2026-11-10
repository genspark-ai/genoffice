import type { Transaction } from '@tiptap/pm/state'
import { ReplaceStep } from '@tiptap/pm/transform'

/**
 * Top-level blocks a transaction touched, as offsets in its final document;
 * null when a step replaces so much that recomputing everything is cheaper.
 */
export function touchedTopLevelBlocks(tr: Transaction): Set<number> | null {
  const { doc } = tr
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
  if (unknown) return null
  const touched = new Set<number>()
  doc.forEach((node, offset) => {
    const end = offset + node.nodeSize
    if (ranges.some(([from, to]) => offset < to && end > from)) touched.add(offset)
  })
  return touched.size * 2 > doc.childCount ? null : touched
}

/** a single insertion at the document end: every existing position survives unchanged */
export function appendsAtEnd(tr: Transaction): boolean {
  if (tr.steps.length !== 1) return false
  const step = tr.steps[0]
  if (!(step instanceof ReplaceStep)) return false
  return step.from === step.to && step.from === tr.before.content.size
}
