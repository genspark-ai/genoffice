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
 */
export class TopLevelPositions {
  private map: Map<Node, { from: number; to: number }> | null = null

  constructor(private readonly view: EditorView) {}

  /** before/after of the top-level block holding `el`; null when the element is not mounted */
  of(el: Element): { from: number; to: number } | null {
    let top: Element | null = el
    while (top && top.parentElement !== this.view.dom) top = top.parentElement
    if (!top) return null
    if (!this.map) {
      this.map = new Map()
      const { children } = (this.view as unknown as { docView: { children: ChildDesc[] } }).docView
      let pos = 0
      for (const child of children) {
        if (child.size > 0) this.map.set(child.dom, { from: pos, to: pos + child.size })
        pos += child.size
      }
    }
    return this.map.get(top) ?? null
  }
}
