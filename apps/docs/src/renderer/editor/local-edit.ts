import type { Node as PmNode } from '@tiptap/pm/model'
import { Plugin, PluginKey, type Transaction } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import { touchedTopLevelBlocks } from './touched-blocks'

/**
 * Block-attribute decoration plugins used to rebuild their whole set on every
 * view update by walking every top-level node — five of them, per keystroke,
 * on a 10k-block document. A local edit (block count unchanged, a few blocks
 * touched) can only change decorations on the blocks it touched: when none of
 * those carries the attribute now and none carried a decoration before, the
 * old set maps through the transaction unchanged.
 */

/** offsets (in tr.doc) of the top-level blocks a local edit touched; null when the edit is not local */
export function localEditBlocks(tr: Transaction): Set<number> | null {
  if (tr.before.childCount !== tr.doc.childCount) return null
  return touchedTopLevelBlocks(tr)
}

/**
 * Whether the mapped old set or the given predicate wants a recompute for any
 * touched block. The predicate sees the block before and after the edit:
 * what a block turned into and what it stopped being both matter (a list
 * item made a paragraph still renumbers the items after it). Block count is
 * unchanged for a local edit, so the same index addresses both documents.
 */
export function touchedNeedsRecompute(
  tr: Transaction,
  touched: Set<number>,
  mapped: DecorationSet,
  needsRecompute: (node: PmNode) => boolean,
): boolean {
  for (const offset of touched) {
    const child = tr.doc.childAfter(offset)
    if (!child.node) return true
    if (needsRecompute(child.node)) return true
    const before = tr.before.maybeChild(child.index)
    if (before && needsRecompute(before)) return true
    if (mapped.find(offset, offset + child.node.nodeSize).length > 0) return true
  }
  return false
}

/** the node itself or any descendant satisfies pred (stops at the first hit) */
export function containsNode(node: PmNode, pred: (n: PmNode) => boolean): boolean {
  if (pred(node)) return true
  let found = false
  node.descendants((n) => {
    if (found) return false
    if (pred(n)) found = true
    return !found
  })
  return found
}

interface BlockDecorationState {
  set: DecorationSet
  sig: string
}

export interface BlockDecorationSpec {
  key: PluginKey<BlockDecorationState>
  /** decorations of the whole document */
  build: (doc: PmNode) => Decoration[]
  /** a touched block whose current content forces a rebuild */
  needsRecompute: (node: PmNode) => boolean
  /** inputs outside the document (display modes): a change rebuilds on the next transaction */
  signature?: () => string
}

export function blockDecorationPlugin(spec: BlockDecorationSpec): Plugin<BlockDecorationState> {
  const compute = (doc: PmNode): BlockDecorationState => {
    const decos = spec.build(doc)
    return {
      set: decos.length > 0 ? DecorationSet.create(doc, decos) : DecorationSet.empty,
      sig: spec.signature?.() ?? '',
    }
  }
  return new Plugin<BlockDecorationState>({
    key: spec.key,
    state: {
      init: (_config, state) => compute(state.doc),
      apply(tr, old) {
        const sig = spec.signature?.() ?? ''
        if (sig !== old.sig) return compute(tr.doc)
        if (!tr.docChanged) return old
        const touched = localEditBlocks(tr)
        if (!touched) return compute(tr.doc)
        const mapped = old.set.map(tr.mapping, tr.doc)
        if (touchedNeedsRecompute(tr, touched, mapped, spec.needsRecompute)) return compute(tr.doc)
        return { set: mapped, sig }
      },
    },
    props: {
      decorations(state) {
        return spec.key.getState(state)?.set ?? null
      },
    },
  })
}
