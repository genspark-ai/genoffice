import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import type { EditorView } from '@tiptap/pm/view'

/** CSS floats: the only layout where a paragraph's vertical position changes its line breaks */
const FLOAT_SELECTOR =
  '[class*="-wrap-square-"], [class*="-wrap-tight-"], [class*="-wrap-through-"], ' +
  '[class*="doc-table-float-"], .doc-cell-boxes'

interface Entry<T> {
  node: ProseMirrorNode
  key: string
  rectKey: string
  result: T
  settled: boolean
}

/**
 * Per-paragraph result cache for the measure → decorate → re-measure loops of
 * the shrink extensions. The loop re-enters synchronously, so only the
 * extension's own decorations change between passes: a paragraph whose last
 * two results agree and whose box is unchanged would measure the same again,
 * and the costly per-word/per-char DOM walk is skipped. Callers clear() on
 * every external input (edit, resize, font load).
 */
export class SettledParagraphCache<T> {
  private results = new Map<number, Entry<T>>()
  private floatBands: number[][] = []

  clear(): void {
    this.results.clear()
  }

  /** once per pass: vertical bands beside which text wraps */
  beginPass(view: EditorView): void {
    this.floatBands = Array.from(view.dom.querySelectorAll(FLOAT_SELECTOR), (f) => {
      const r = f.getBoundingClientRect()
      const cs = getComputedStyle(f)
      return [
        r.top - (parseFloat(cs.marginTop) || 0),
        r.bottom + (parseFloat(cs.marginBottom) || 0),
      ]
    })
  }

  measure(
    view: EditorView,
    node: ProseMirrorNode,
    pos: number,
    measureParagraph: () => T | null,
  ): T | null {
    const el = view.nodeDOM(pos)
    if (!(el instanceof HTMLElement)) return null
    const r = el.getBoundingClientRect()
    const nearFloat = this.floatBands.some(([top, bottom]) => r.top < bottom && r.bottom > top)
    const rectKey = `${nearFloat ? r.top : 0}:${r.height}:${r.width}`
    const prev = this.results.get(pos)
    if (prev?.settled && prev.node === node && prev.rectKey === rectKey) return prev.result
    const result = measureParagraph()
    if (result === null) {
      this.results.delete(pos)
      return null
    }
    const key = JSON.stringify(result)
    this.results.set(pos, {
      node,
      key,
      rectKey,
      result,
      settled: prev?.node === node && prev.key === key,
    })
    return result
  }
}
