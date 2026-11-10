import { DecorationSet } from '@tiptap/pm/view'
import type { Decoration } from '@tiptap/pm/view'
import { Fragment } from '@tiptap/pm/model'
import type { Node as PmNode } from '@tiptap/pm/model'

/**
 * prosemirror-view's DecorationSet.forChild scans every root-level decoration
 * and every child entry for each child of a node. Long documents keep one
 * node decoration per paragraph at the root (line factors, anchors), so each
 * view update costs blocks × decorations — hundreds of ms per keystroke on a
 * 10k-paragraph document. Non-inline root decorations can never apply to a
 * child, and the child table is sorted, so skip the former and binary-search
 * the latter. Same result as upstream; the mixed case (inline root
 * decorations overlapping a child that has its own subtree) still goes
 * through the original method, which builds the private DecorationGroup.
 */
interface SetInternals {
  local: Decoration[]
  children: Array<number | SetInternals>
  inlineLocals?: Decoration[]
}

type ForChild = (this: DecorationSet, offset: number, node: PmNode) => DecorationSet

const proto = DecorationSet.prototype as unknown as { forChild: ForChild }
const original = proto.forChild
const Ctor = DecorationSet as unknown as new (
  local: Decoration[],
  children: Array<number | SetInternals>,
) => DecorationSet
const isInline = (d: Decoration) => (d as unknown as { inline: boolean }).inline
const byPos = (a: Decoration, b: Decoration) => a.from - b.from || a.to - b.to

export function fastForChild(this: DecorationSet, offset: number, node: PmNode): DecorationSet {
  if (this === DecorationSet.empty) return this
  if (node.isLeaf) return DecorationSet.empty
  const self = this as unknown as SetInternals
  const children = self.children
  let lo = 0
  let hi = children.length / 3
  while (lo < hi) {
    const mid = (lo + hi) >> 1
    if ((children[mid * 3] as number) < offset) lo = mid + 1
    else hi = mid
  }
  const child =
    lo * 3 < children.length && children[lo * 3] === offset
      ? (children[lo * 3 + 2] as unknown as DecorationSet)
      : undefined
  const inline = self.inlineLocals ?? (self.inlineLocals = self.local.filter(isInline))
  if (inline.length === 0) return child ?? DecorationSet.empty
  if (child) return original.call(this, offset, node)
  const start = offset + 1
  const end = start + node.content.size
  let local: Decoration[] | undefined
  for (const dec of inline) {
    if (dec.from < end && dec.to > start) {
      const from = Math.max(start, dec.from) - start
      const to = Math.min(end, dec.to) - start
      if (from < to)
        (local ??= []).push(
          (dec as unknown as { copy: (f: number, t: number) => Decoration }).copy(from, to),
        )
    }
  }
  return local ? new Ctor(local.sort(byPos), []) : DecorationSet.empty
}

/**
 * Fragment.findIndex walks the children from the front on every call. Node
 * decorations are validated through it on every DecorationSet.map, so a
 * keystroke on a 10k-paragraph document cost paragraphs² again. Fragments are
 * immutable: cache the children's end positions once and binary-search.
 */
interface FragmentInternals {
  content: PmNode[]
  size: number
  ends?: Float64Array
}
type FindIndex = (this: Fragment, pos: number) => { index: number; offset: number }
const fragmentProto = Fragment.prototype as unknown as { findIndex: FindIndex }
const originalFindIndex = fragmentProto.findIndex
/** below this the linear walk wins */
const FIND_INDEX_MIN_CHILDREN = 64
const found = { index: 0, offset: 0 }

export function fastFindIndex(this: Fragment, pos: number): { index: number; offset: number } {
  const self = this as unknown as FragmentInternals
  const content = self.content
  if (content.length < FIND_INDEX_MIN_CHILDREN || pos <= 0 || pos >= self.size) {
    return originalFindIndex.call(this, pos)
  }
  let ends = self.ends
  if (!ends) {
    ends = new Float64Array(content.length)
    let acc = 0
    for (let i = 0; i < content.length; i++) ends[i] = acc += content[i].nodeSize
    self.ends = ends
  }
  let lo = 0
  let hi = content.length - 1
  while (lo < hi) {
    const mid = (lo + hi) >> 1
    if (ends[mid] < pos) lo = mid + 1
    else hi = mid
  }
  if (ends[lo] === pos) {
    found.index = lo + 1
    found.offset = pos
  } else {
    found.index = lo
    found.offset = lo === 0 ? 0 : ends[lo - 1]
  }
  return found
}

export function installProseMirrorPerf(): void {
  if (proto.forChild !== fastForChild) proto.forChild = fastForChild
  if (fragmentProto.findIndex !== fastFindIndex) fragmentProto.findIndex = fastFindIndex
}

/** test hook: run a callback against the upstream implementations */
export function withOriginalFindIndex<T>(fn: () => T): T {
  fragmentProto.findIndex = originalFindIndex
  try {
    return fn()
  } finally {
    fragmentProto.findIndex = fastFindIndex
  }
}

/** test hook: run a callback against the upstream implementation */
export function withOriginalForChild<T>(fn: () => T): T {
  proto.forChild = original
  try {
    return fn()
  } finally {
    proto.forChild = fastForChild
  }
}
