import type { Node as PmNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type DecorationAttrs } from '@tiptap/pm/view'

interface LayoutAttrs extends DecorationAttrs {
  class: string
  style: string
}

interface ParagraphLayout {
  attrs: LayoutAttrs
  continuation: LayoutAttrs
}

/** The producer explicitly opts in only layouts safe to carry across a split. */
export function columnDecoration(
  from: number,
  to: number,
  attrs: LayoutAttrs,
  continuation?: LayoutAttrs,
): Decoration {
  return Decoration.node(from, to, attrs, {
    key: `col-${attrs.class}-${attrs.style}`,
    ...(continuation ? { paragraphLayout: { attrs, continuation } satisfies ParagraphLayout } : {}),
  })
}

/**
 * Enter maps a paragraph's node decoration across two nodes, invalidating it.
 * Recover only the opted-in display frame, owned by the original first block:
 * an undecorated first paragraph must stay neutral when joined with a decorated
 * second one. Leading-only flags stay on the first resulting paragraph.
 * No model attrs, document edits, DOM measurements or extra transactions.
 */
export function mapColumnDecorations(
  previous: DecorationSet,
  tr: Transaction,
  oldDoc: PmNode,
): DecorationSet {
  let removed: Set<unknown> | null = null
  const mapped = previous.map(tr.mapping, tr.doc, {
    onRemove: (spec) => (removed ??= new Set()).add(spec),
  })
  if (!removed) return mapped
  const dropped = previous.find(
    undefined,
    undefined,
    (spec) => removed!.has(spec) && spec.paragraphLayout,
  )
  if (dropped.length === 0) return mapped

  const roots: Array<{ from: number; to: number; textblock: boolean }> = []
  const textblocks = new Set<number>()
  oldDoc.forEach((node, from) => {
    roots.push({ from, to: from + node.nodeSize, textblock: node.isTextblock })
    if (node.isTextblock) textblocks.add(from)
  })
  const ownerAt = (pos: number): number | null => {
    let lo = 0
    let hi = roots.length
    while (lo < hi) {
      const mid = (lo + hi) >>> 1
      if (roots[mid].from <= pos) lo = mid + 1
      else hi = mid
    }
    const root = roots[lo - 1]
    return root?.textblock && pos < root.to ? root.from : null
  }

  const recovery = new Map<number, { from: number; to: number; layout: ParagraphLayout }>()
  let first = Infinity
  let last = -Infinity
  for (const d of dropped) {
    if (!textblocks.has(d.from)) continue
    const from = tr.mapping.map(d.from, 1)
    const to = tr.mapping.map(d.to, -1)
    // Deleted/replaced whole blocks must not lend their layout to replacements.
    if (from >= to) continue
    recovery.set(d.from, { from, to, layout: d.spec.paragraphLayout as ParagraphLayout })
    first = Math.min(first, from)
    last = Math.max(last, to)
  }
  if (recovery.size === 0) return mapped

  // One scan rather than a DecorationSet.find per restored block: root-level
  // node decorations share a flat list, so repeated finds make bulk edits quadratic.
  const occupied = new Set(
    mapped
      .find()
      .filter((d) => d.from < d.to)
      .map((d) => d.from),
  )
  const inverse = tr.mapping.invert()
  const used = new Set<number>()
  const carried: Decoration[] = []
  tr.doc.forEach((node, pos) => {
    const end = pos + node.nodeSize
    if (!node.isTextblock || occupied.has(pos) || end <= first || pos >= last) return
    // The outer start, rather than the first character, preserves the first
    // paragraph's ownership even when it was empty before a join.
    const owner = ownerAt(inverse.map(pos, -1))
    if (owner === null) return
    const entry = recovery.get(owner)
    if (!entry || end <= entry.from || pos >= entry.to) return
    const attrs = used.has(owner) ? entry.layout.continuation : entry.layout.attrs
    carried.push(columnDecoration(pos, end, attrs, entry.layout.continuation))
    used.add(owner)
  })
  return carried.length > 0 ? mapped.add(tr.doc, carried) : mapped
}
