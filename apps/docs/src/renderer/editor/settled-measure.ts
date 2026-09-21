import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import { AddMarkStep, RemoveMarkStep, ReplaceStep } from '@tiptap/pm/transform'
import type { EditorView } from '@tiptap/pm/view'

/** CSS floats: the only layout where a paragraph's vertical position changes its line breaks */
// wrap-around floats the paragraphs flow beside (explicit classes: an
// attribute-substring selector scanned the whole editor on every keystroke)
const FLOAT_SELECTOR = [
  ...['square', 'tight', 'through'].flatMap((mode) =>
    ['left', 'right'].flatMap((side) => [
      `.img-wrap-${mode}-${side}`,
      `.doc-inline-img--wrap-${mode}-${side}`,
    ]),
  ),
  '.doc-table-float-left',
  '.doc-table-float-right',
  '.doc-table-float-center',
  '.doc-cell-boxes',
].join(', ')

/** float bands of one editor state, shared by every cache measuring against it */
const floatBandsByState = new WeakMap<object, number[][]>()
/** the float elements of the last scan of one editor; null once a transaction may have added one */
let floatScan: { dom: Element; els: Element[] } | null = null

/** Only a transaction that inserts something other than text can add a float:
 *  plain typing keeps the element list and re-reads the boxes. */
export function noteFloatTransaction(tr: Transaction): void {
  if (!floatScan || !tr.docChanged) return
  for (const step of tr.steps) {
    if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) continue
    if (step instanceof ReplaceStep && step.slice.openStart === 0 && step.slice.openEnd === 0) {
      let text = true
      step.slice.content.forEach((n) => {
        if (!n.isText) text = false
      })
      if (text) continue
    }
    floatScan = null
    return
  }
}

function floatBandsOf(view: EditorView): number[][] {
  const state = (view as { state?: object }).state
  const cached = state && floatBandsByState.get(state)
  if (cached) return cached
  if (!floatScan || floatScan.dom !== view.dom || floatScan.els.some((f) => !f.isConnected))
    floatScan = { dom: view.dom, els: Array.from(view.dom.querySelectorAll(FLOAT_SELECTOR)) }
  // a floating table the paginator flows in place (doc-table-float-flow, a
  // decoration class) is not a float for this pass
  const bands = floatScan.els
    .filter((f) => !f.classList.contains('doc-table-float-flow'))
    .map((f) => {
      const r = f.getBoundingClientRect()
      const cs = getComputedStyle(f)
      return [
        r.top - (parseFloat(cs.marginTop) || 0),
        r.bottom + (parseFloat(cs.marginBottom) || 0),
      ]
    })
  if (state) floatBandsByState.set(state, bands)
  return bands
}

interface Entry<T> {
  /** paragraph start when measured; results are stored at that offset */
  pos: number
  /** position-independent identity of the result */
  key: string
  rectKey: string
  result: T
  settled: boolean
  /** last pass that visited the paragraph (unvisited entries are pruned) */
  gen: number
}

/**
 * Per-paragraph result cache for the measure → decorate → re-measure loops of
 * the shrink extensions. The loop re-enters synchronously, so only the
 * extension's own decorations change between passes: a paragraph whose last
 * two results agree and whose box is unchanged would measure the same again,
 * and the costly per-word/per-char DOM walk is skipped.
 *
 * Entries key on the paragraph node's identity, so an edit elsewhere in the
 * document keeps every untouched paragraph's settled result (ProseMirror
 * shares unchanged nodes across states); only its positions shift, which
 * `shift` re-bases. Callers clear() when layout inputs change (resize, font
 * load, document stylesheet).
 */
export class SettledParagraphCache<T> {
  private results = new Map<ProseMirrorNode, Entry<T>>()
  private floatBands: number[][] = []
  private gen = 0

  /**
   * @param shift re-bases a stored result to a paragraph that moved by delta
   * @param stableNow a result that changes nothing about the paragraph's
   *   own decorations (nothing to shrink) cannot alter its layout, so it is
   *   settled after a single pass
   */
  constructor(
    private shift: (result: T, delta: number) => T,
    private stableNow: (result: T) => boolean = () => false,
  ) {}

  clear(): void {
    this.results.clear()
  }

  /** once per pass: vertical bands beside which text wraps; drops paragraphs the previous pass never visited */
  beginPass(view: EditorView): void {
    this.gen++
    for (const [node, entry] of this.results)
      if (entry.gen < this.gen - 1) this.results.delete(node)
    this.floatBands = floatBandsOf(view)
  }

  /**
   * Top-level block nodes → their DOM, read off ProseMirror's view descriptors
   * in one walk; view.nodeDOM(pos) descends through every top-level block for
   * each call, which made a pass over N paragraphs cost N².
   */
  static topLevelDom(view: EditorView): Map<ProseMirrorNode, HTMLElement> {
    const map = new Map<ProseMirrorNode, HTMLElement>()
    for (const child of Array.from(view.dom.childNodes)) {
      const desc = (child as { pmViewDesc?: { node?: ProseMirrorNode } }).pmViewDesc
      if (desc?.node && child instanceof HTMLElement) map.set(desc.node, child)
    }
    return map
  }

  measure(
    view: EditorView,
    node: ProseMirrorNode,
    pos: number,
    measureParagraph: (el: HTMLElement) => T | null,
    dom?: HTMLElement,
  ): T | null {
    const el = dom ?? view.nodeDOM(pos)
    if (!(el instanceof HTMLElement)) return null
    // layout px: a page-zoom transform changes no line break
    let rectKey = `${el.offsetHeight}:${el.offsetWidth}`
    if (this.floatBands.length > 0) {
      const r = el.getBoundingClientRect()
      if (this.floatBands.some(([top, bottom]) => r.top < bottom && r.bottom > top))
        rectKey = `${r.top}:${rectKey}`
    }
    const prev = this.results.get(node)
    if (prev) prev.gen = this.gen
    if (prev?.settled && prev.rectKey === rectKey) {
      return prev.pos === pos ? prev.result : this.shift(prev.result, pos - prev.pos)
    }
    const result = measureParagraph(el)
    if (result === null) {
      this.results.delete(node)
      return null
    }
    const key = JSON.stringify(this.shift(result, -pos))
    this.results.set(node, {
      pos,
      key,
      rectKey,
      result,
      settled: (prev !== undefined && prev.key === key) || this.stableNow(result),
      gen: this.gen,
    })
    return result
  }
}
