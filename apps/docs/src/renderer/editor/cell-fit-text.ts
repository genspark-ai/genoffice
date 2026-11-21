/**
 * w:tcFitText (Table Properties > Cell > Options > Fit text): every paragraph
 * of the cell is one line as wide as the cell. A short line is spread between
 * its characters (CSS justification); a long one is squeezed, and how far is
 * only known once it is laid out, so the squeeze is a display-only decoration
 * measured from the DOM.
 */
import { Extension } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import { DOC_CSS_COMMITTED_EVENT, WRAP_WIDTHS_COMMITTED_EVENT } from './cjk-punct-shrink'

interface Squeeze {
  /** horizontal scale, below 1 */
  k: number
  /** px from the paragraph's left edge to where its text starts: the point that stays put */
  origin: number
}

/** by paragraph position */
type Squeezes = ReadonlyMap<number, Squeeze>

export const cellFitTextKey = new PluginKey<Squeezes>('cellFitText')

const CELLS = new Set(['docTableCell', 'docTableHeader'])
const fitLines = new WeakMap<ProseMirrorNode, number[]>()

/** the paragraphs of the document's fit-text cells; most documents have none */
export function fitTextLines(doc: ProseMirrorNode): number[] {
  let lines = fitLines.get(doc)
  if (!lines) {
    const found: number[] = []
    doc.descendants((node, pos) => {
      if (node.isTextblock) return false
      if (CELLS.has(node.type.name) && node.attrs.fitText)
        node.forEach((child, offset) => {
          if (child.isTextblock) found.push(pos + 1 + offset)
        })
      return true
    })
    lines = found
    fitLines.set(doc, lines)
  }
  return lines
}

/** how far a line of `natural` width is squeezed to fit `avail`; 1 = not at all */
export function fitTextScale(natural: number, avail: number): number {
  if (!(natural > 0) || !(avail > 0) || natural <= avail + 0.5) return 1
  return Math.round((avail / natural) * 1000) / 1000
}

class CellFitTextView {
  private raf = 0
  private resizeObserver?: ResizeObserver
  private width = 0

  constructor(private view: EditorView) {
    document.addEventListener(DOC_CSS_COMMITTED_EVENT, this.schedule)
    document.addEventListener(WRAP_WIDTHS_COMMITTED_EVENT, this.schedule)
    document.fonts?.addEventListener('loadingdone', this.schedule)
    if (typeof ResizeObserver !== 'undefined') {
      this.resizeObserver = new ResizeObserver(() => {
        if (view.dom.offsetWidth === this.width) return
        this.width = view.dom.offsetWidth
        this.schedule()
      })
      this.resizeObserver.observe(view.dom)
    }
    this.schedule()
  }

  update() {
    this.schedule()
  }

  destroy() {
    document.removeEventListener(DOC_CSS_COMMITTED_EVENT, this.schedule)
    document.removeEventListener(WRAP_WIDTHS_COMMITTED_EVENT, this.schedule)
    document.fonts?.removeEventListener('loadingdone', this.schedule)
    this.resizeObserver?.disconnect()
    if (this.raf) cancelAnimationFrame(this.raf)
  }

  private schedule = () => {
    if (this.raf) return
    this.raf = requestAnimationFrame(() => {
      this.raf = 0
      this.measure()
    })
  }

  private measure() {
    const { view } = this
    if (!view.dom.isConnected) return
    const lines = fitTextLines(view.state.doc)
    const applied = cellFitTextKey.getState(view.state) ?? new Map<number, Squeeze>()
    if (lines.length === 0 && applied.size === 0) return
    // PDF export parks the editor subtree: a layout read would lay it out again per print chunk
    if (view.dom.closest('.app.pv-exporting')) return
    const next = new Map<number, Squeeze>()
    let changed = false
    for (const pos of lines) {
      const had = applied.get(pos)
      const now = this.squeezeOf(pos, had?.k ?? 1)
      // a paragraph that is not laid out right now keeps the squeeze it has
      if (now === null) {
        if (had) next.set(pos, had)
        continue
      }
      // a re-measure of a squeezed line lands a hair off: that is not a change
      const same = had
        ? Math.abs(now.k - had.k) < 0.004 && Math.abs(now.origin - had.origin) < 0.5
        : now.k === 1
      if (!same) changed = true
      const kept = same ? had : now
      if (kept && kept.k < 1) next.set(pos, kept)
    }
    if (!changed && next.size === applied.size) return
    view.dispatch(view.state.tr.setMeta(cellFitTextKey, next).setMeta('addToHistory', false))
  }

  /** null = not laid out right now (hidden, not mounted) */
  private squeezeOf(pos: number, applied: number): Squeeze | null {
    const el = this.view.nodeDOM(pos)
    if (!(el instanceof HTMLElement) || el.offsetWidth === 0) return null
    const cs = getComputedStyle(el)
    const px = (v: string) => parseFloat(v) || 0
    const avail =
      el.clientWidth - px(cs.paddingLeft) - px(cs.paddingRight) - Math.max(0, px(cs.textIndent))
    // the box is drawn squeezed and zoomed; the text's own width is wanted
    const drawn = el.getBoundingClientRect().width / (el.offsetWidth * applied) || 1
    const range = document.createRange()
    range.selectNodeContents(el)
    const natural = range.getBoundingClientRect().width / (drawn * applied)
    return {
      k: fitTextScale(natural, avail),
      origin: Math.round((px(cs.borderLeftWidth) + px(cs.paddingLeft)) * 10) / 10,
    }
  }
}

export const CellFitTextExtension = Extension.create({
  name: 'cellFitText',
  addProseMirrorPlugins() {
    return [
      new Plugin<Squeezes>({
        key: cellFitTextKey,
        state: {
          init: () => new Map(),
          apply(tr, old) {
            const meta = tr.getMeta(cellFitTextKey) as Squeezes | undefined
            if (meta) return meta
            if (!tr.docChanged || old.size === 0) return old
            const mapped = new Map<number, Squeeze>()
            for (const [pos, squeeze] of old) mapped.set(tr.mapping.map(pos), squeeze)
            return mapped
          },
        },
        props: {
          decorations(state) {
            const lines = fitTextLines(state.doc)
            if (lines.length === 0) return null
            const squeezes = cellFitTextKey.getState(state)
            return DecorationSet.create(
              state.doc,
              lines.map((pos) => {
                const squeeze = squeezes?.get(pos)
                const node = state.doc.nodeAt(pos)!
                return Decoration.node(
                  pos,
                  pos + node.nodeSize,
                  squeeze
                    ? {
                        class: 'doc-fit-line doc-fit-squeezed',
                        style: `transform:scaleX(${squeeze.k});transform-origin:${squeeze.origin}px 0`,
                      }
                    : { class: 'doc-fit-line' },
                )
              }),
            )
          },
        },
        view: (view) => new CellFitTextView(view),
      }),
    ]
  },
})
