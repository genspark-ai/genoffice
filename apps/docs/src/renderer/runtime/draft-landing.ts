import type { Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import { TABLE_TRAILING_SKIP } from '../editor/extensions'
import { TRACK_IGNORE } from '../editor/revisions'
import { blockRangePositions, parseHtmlFragment, type NumIds } from '../ai/protocol'

export const DOC_MAX_CHARS = 200_000
const DRAFT_TICK_MS = 300

export type WritePosition = { kind: 'whole' } | { kind: 'after'; index: number }

/** meta on every draft transaction: not an edit, not history, not a tracked change */
export const DRAFT_META = 'aiDraft'
const activeDrafts = new WeakMap<Editor, number>()

export function hasActiveDocumentDraft(editor: Editor): boolean {
  return (activeDrafts.get(editor) ?? 0) > 0
}

/**
 * Live draft of a streaming fragment inside the document. Each tick re-parses
 * the fragment received so far (the HTML parser closes open tags, so a
 * half-received block still renders) and swaps only the blocks that changed.
 * The draft is outside undo history and Track Changes; `finish()` removes it
 * so the caller can land the final fragment through the regular insert path,
 * giving one undo step and normal revision marks.
 */
export class DraftLanding {
  private from: number
  private to: number
  /** the text of the last tick that parsed and rendered */
  private rendered: string | null = null
  private pending: { html: string; nodes?: ReturnType<typeof parseHtmlFragment> } | null = null
  private timer: ReturnType<typeof setTimeout> | null = null
  private closed = false
  private follow = true
  private conflicted = false
  private ownedNodes = new Set<ProseMirrorNode>()
  get hasConflict(): boolean {
    return this.conflicted
  }
  private readonly onTransaction = ({ transaction }: { transaction: Transaction }) => {
    if (!transaction.docChanged || transaction.getMeta(DRAFT_META)) return
    for (const map of transaction.mapping.maps) {
      if (this.to > this.from)
        map.forEach((start, end) => {
          if (
            (start < this.to && end > this.from) ||
            (start === end && start > this.from && start < this.to)
          )
            this.conflicted = true
        })
      this.from = map.map(this.from, 1)
      this.to = Math.max(this.from, map.map(this.to, -1))
    }
  }

  constructor(
    private readonly editor: Editor,
    private readonly numIds: NumIds,
    position: WritePosition,
    private readonly canPreview: () => boolean = () => !editor.isDestroyed && editor.isEditable,
  ) {
    this.from =
      position.kind === 'after' && position.index >= 0
        ? blockRangePositions(editor, position.index, position.index).to
        : 0
    this.to = this.from
    activeDrafts.set(editor, (activeDrafts.get(editor) ?? 0) + 1)
    editor.on('transaction', this.onTransaction)
  }

  /** latest cumulative fragment; rendered on the next tick */
  update(html: string, nodes?: ReturnType<typeof parseHtmlFragment>): void {
    if (this.closed) return
    this.pending = { html, nodes }
    if (this.timer === null) this.timer = setTimeout(() => this.flush(), DRAFT_TICK_MS)
  }

  /** the top-level block index the draft sits after (-1 = document start), after any user edits */
  indexBefore(): number {
    const doc = this.editor.state.doc
    let index = -1
    let i = 0
    doc.forEach((node, offset) => {
      if (offset + node.nodeSize <= this.from) index = i
      i++
    })
    return index
  }

  /** Freeze the visible preview without publishing a pending stale update. */
  pause(): void {
    if (this.timer !== null) clearTimeout(this.timer)
    this.timer = null
    this.pending = null
  }

  /**
   * Remove the draft; the document is back to what it was around it. Returns the
   * text of the last tick that rendered, so a kept stream whose tail no longer
   * parses (a half-received formula or table) can still land what the user saw.
   */
  finish(remove = true): string | null {
    if (this.closed) return this.rendered
    if (this.timer !== null) clearTimeout(this.timer)
    this.timer = null
    if (remove) {
      this.conflicted ||= this.liveNodes().some(({ node }) => !this.ownedNodes.has(node))
      if (!this.conflicted) this.flush()
    }
    this.closed = true
    activeDrafts.set(this.editor, Math.max(0, (activeDrafts.get(this.editor) ?? 1) - 1))
    this.editor.off('transaction', this.onTransaction)
    if (remove && this.to > this.from && !this.editor.isDestroyed) {
      let tr = this.draftTr()
      if (this.conflicted) {
        // Editing a preview block adopts it as user content. Remove only the
        // untouched draft blocks, preserving those edits and their history.
        for (const { node, end } of this.liveNodes().reverse())
          if (this.ownedNodes.has(node)) tr = tr.delete(end - node.nodeSize, end)
      } else tr = tr.delete(this.from, this.to)
      if (tr.docChanged) this.editor.view.dispatch(tr)
    }
    this.to = this.from
    return this.rendered
  }

  /** live top-level blocks inside the draft range, after any user edits */
  private liveNodes(): Array<{ node: ProseMirrorNode; end: number }> {
    const out: Array<{ node: ProseMirrorNode; end: number }> = []
    this.editor.state.doc.forEach((node, offset) => {
      if (offset >= this.from && offset + node.nodeSize <= this.to)
        out.push({ node, end: offset + node.nodeSize })
    })
    return out
  }

  private draftTr(): Transaction {
    return this.editor.state.tr
      .setMeta(DRAFT_META, true)
      .setMeta('addToHistory', false)
      .setMeta(TRACK_IGNORE, true)
      .setMeta(TABLE_TRAILING_SKIP, true)
  }

  private flush(): void {
    this.timer = null
    if (this.closed || this.conflicted || this.pending === null || !this.canPreview()) return
    const { html, nodes } = this.pending
    this.pending = null
    let next: ProseMirrorNode[]
    try {
      next = (nodes ?? parseHtmlFragment(html, this.numIds)).map((n) =>
        this.editor.schema.nodeFromJSON({ ...n, attrs: { ...n.attrs, aiChanged: false } }),
      )
    } catch {
      return // a half-received formula or table is retried on the next tick
    }
    const live = this.liveNodes()
    if (live.some(({ node }) => !this.ownedNodes.has(node))) {
      this.conflicted = true
      return
    }
    let keep = 0
    while (keep < live.length && keep < next.length && live[keep]!.node.eq(next[keep]!)) keep++
    if (keep === live.length && keep === next.length) return
    const at = keep > 0 ? live[keep - 1]!.end : this.from
    const wasVisible = this.follow && this.endVisible()
    if (!this.canPreview()) return
    this.rendered = html
    this.editor.view.dispatch(this.draftTr().replaceWith(at, this.to, next.slice(keep)))
    this.to = at + next.slice(keep).reduce((size, n) => size + n.nodeSize, 0)
    this.ownedNodes = new Set(this.liveNodes().map(({ node }) => node))
    if (wasVisible) this.scrollToEnd()
    else this.follow = false
  }

  private endVisible(): boolean {
    if (this.to === this.from) return true
    try {
      const { top } = this.editor.view.coordsAtPos(this.to - 1)
      return top >= 0 && top <= window.innerHeight
    } catch {
      return true
    }
  }

  private scrollToEnd(): void {
    if (this.to === this.from) return
    try {
      const { node } = this.editor.view.domAtPos(this.to - 1)
      const el = node instanceof Element ? node : node.parentElement
      el?.scrollIntoView?.({ block: 'nearest' })
    } catch {
      /* layout not ready */
    }
  }
}
