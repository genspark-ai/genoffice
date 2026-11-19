import type { Node as PmNode } from '@tiptap/pm/model'
import type { EditorView } from '@tiptap/pm/view'

interface ChildDesc {
  dom: Node
  size: number
}

/**
 * Doc positions of the top-level blocks from one walk over the view's child
 * descs. posAtDOM sums the sizes of every preceding sibling on each call, so
 * a pass positioning thousands of blocks (column specs, float shifts, gaps)
 * paid blocks² for it — tens of seconds on a long converted PDF.
 *
 * The walk is cached per document: view.state.doc is a new object after every
 * doc-changing transaction, so a changed identity means the child descs (and
 * their sizes) may have moved. invalidate() forces a rebuild for callers that
 * mutate DOM without a doc change. A posAtDOM fallback covers elements the
 * walk missed (stale docView children); widgets resolve to a neighbouring
 * block, so the fallback only accepts when nodeDOM confirms ownership.
 */
export class TopLevelPositions {
  private map: Map<Node, { from: number; to: number }> | null = null
  private doc: PmNode | null = null

  constructor(private readonly view: EditorView) {}

  /** drop the cached walk; the next of() rebuilds from the current docView */
  invalidate(): void {
    this.map = null
    this.doc = null
  }

  /** before/after of the top-level block holding `el`; null when the element is not mounted */
  of(el: Element): { from: number; to: number } | null {
    let top: Element | null = el
    while (top && top.parentElement !== this.view.dom) top = top.parentElement
    if (!top) return null
    const cur = this.view.state.doc
    if (!this.map || this.doc !== cur) {
      this.map = new Map()
      const { children } = (this.view as unknown as { docView: { children: ChildDesc[] } }).docView
      let pos = 0
      for (const child of children) {
        if (child.size > 0) this.map.set(child.dom, { from: pos, to: pos + child.size })
        pos += child.size
      }
      this.doc = cur
    }
    const hit = this.map.get(top)
    if (hit) return hit
    try {
      const $inside = cur.resolve(this.view.posAtDOM(top, 0))
      if ($inside.depth < 1) return null
      const from = $inside.before(1)
      const to = $inside.after(1)
      if (this.view.nodeDOM(from) !== top) return null
      return { from, to }
    } catch {
      return null
    }
  }
}
