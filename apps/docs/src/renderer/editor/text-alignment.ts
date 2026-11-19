/**
 * w:textAlignment (Paragraph > Asian Typography > Text Alignment): how runs of
 * different sizes stand on one line. Word (probe 2026-11-18, 9 fonts, 5 size
 * pairs) keeps the tallest run where it is and moves each other run by its
 * font's hhea ascent a and descent d at its size:
 *   top     up by  aBig - a
 *   center  up by  ((aBig - dBig) - (a - d)) / 2
 *   bottom  down by kBig * dBig - k * d   (k = 0.75 for an East Asian font, 0.5 otherwise)
 * The line pitch does not change for one font family, so the move is a
 * relative offset that never re-wraps or re-paginates. A line's tallest run is
 * only known once the paragraph is laid out: the offsets are display-only
 * decorations measured from the DOM.
 */
import { Extension } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view'
import { DOC_CSS_COMMITTED_EVENT, WRAP_WIDTHS_COMMITTED_EVENT } from './cjk-punct-shrink'

export type TextAlignmentMode = 'top' | 'center' | 'bottom'

export interface RunBox {
  /** font size, px */
  size: number
  /** hhea ascent and descent at that size, px */
  ascent: number
  descent: number
  eastAsian: boolean
}

/** how far up (px, negative = down) a run moves against the line's tallest run */
export function textAlignmentShift(mode: TextAlignmentMode, big: RunBox, run: RunBox): number {
  if (mode === 'top') return big.ascent - run.ascent
  if (mode === 'center') return (big.ascent - big.descent - (run.ascent - run.descent)) / 2
  const k = (box: RunBox) => (box.eastAsian ? 0.75 : 0.5)
  return -(k(big) * big.descent - k(run) * run.descent)
}

const EAST_ASIAN_FACE =
  /sim(sun|hei)|song|hei|kai|fang|ming|mincho|gothic|yahei|jhenghei|pingfang|hiragino|yu ?(mincho|gothic)|meiryo|malgun|batang|gulim|dotum|noto (sans|serif) (cjk|sc|tc|jp|kr)|source han/i

export const textAlignmentKey = new PluginKey<DecorationSet>('textAlignment')

interface Shift {
  from: number
  to: number
  up: number
}

const MODES = new Set(['top', 'center', 'bottom'])

class TextAlignmentView {
  private raf = 0
  private lastSig = '[]'
  private canvas = document.createElement('canvas').getContext('2d')
  private boxes = new Map<string, { ascent: number; descent: number }>()
  /** whether a document has any paragraph to align: most have none, and then nothing is measured */
  private aligned = new WeakMap<ProseMirrorNode, boolean>()
  private resizeObserver?: ResizeObserver
  private width = 0

  constructor(private view: EditorView) {
    document.addEventListener(DOC_CSS_COMMITTED_EVENT, this.schedule)
    document.addEventListener(WRAP_WIDTHS_COMMITTED_EVENT, this.schedule)
    document.fonts?.addEventListener('loadingdone', this.onFonts)
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

  // any transaction may move a character to another line (typing, a decoration
  // of another pass, the kinsoku level): the pass is cheap when nothing is aligned
  update() {
    this.schedule()
  }

  destroy() {
    document.removeEventListener(DOC_CSS_COMMITTED_EVENT, this.schedule)
    document.removeEventListener(WRAP_WIDTHS_COMMITTED_EVENT, this.schedule)
    document.fonts?.removeEventListener('loadingdone', this.onFonts)
    this.resizeObserver?.disconnect()
    if (this.raf) cancelAnimationFrame(this.raf)
  }

  // the metrics cached so far may be a fallback face's
  private onFonts = () => {
    this.boxes.clear()
    this.schedule()
  }

  private hasAligned(doc: ProseMirrorNode): boolean {
    let has = this.aligned.get(doc)
    if (has === undefined) {
      has = false
      doc.descendants((node) => {
        if (has) return false
        if (!node.isTextblock) return true
        if (MODES.has(node.attrs.textAlignment as string)) has = true
        return false
      })
      this.aligned.set(doc, has)
    }
    return has
  }

  private schedule = () => {
    if (this.raf) return
    this.raf = requestAnimationFrame(() => {
      this.raf = 0
      this.measure()
    })
  }

  private boxOf(font: string): { ascent: number; descent: number } {
    let box = this.boxes.get(font)
    if (!box) {
      box = { ascent: 0, descent: 0 }
      if (this.canvas) {
        this.canvas.font = font
        const m = this.canvas.measureText('x')
        box = { ascent: m.fontBoundingBoxAscent, descent: m.fontBoundingBoxDescent }
      }
      this.boxes.set(font, box)
    }
    return box
  }

  private measure() {
    const { view } = this
    if (!view.dom.isConnected) return
    if (!this.hasAligned(view.state.doc) && this.lastSig === '[]') return
    // PDF export parks the editor subtree: a layout read would lay it out again per print chunk
    if (view.dom.closest('.app.pv-exporting')) return
    const shifts: Shift[] = []
    let unmeasured = false
    view.state.doc.descendants((node, pos) => {
      if (!node.isTextblock) return true
      const mode = node.attrs.textAlignment as string | null
      if (mode && MODES.has(mode)) {
        const measured = this.measureParagraph(node, pos, mode as TextAlignmentMode)
        if (measured) shifts.push(...measured)
        else unmeasured = true
      }
      return false
    })
    // a paragraph that is not laid out right now keeps the offsets it has
    if (unmeasured) return
    const sig = JSON.stringify(shifts)
    if (sig === this.lastSig) return
    this.lastSig = sig
    const decos = shifts.map((s) =>
      Decoration.inline(s.from, s.to, {
        class: 'doc-ta-shift',
        style: `position:relative;top:${-s.up}px`,
      }),
    )
    view.dispatch(view.state.tr.setMeta(textAlignmentKey, decos).setMeta('addToHistory', false))
  }

  /** null = not laid out right now (hidden, not mounted) */
  private measureParagraph(
    node: ProseMirrorNode,
    pos: number,
    mode: TextAlignmentMode,
  ): Shift[] | null {
    const { view } = this
    const para = view.nodeDOM(pos)
    if (!(para instanceof HTMLElement) || para.offsetWidth === 0) return null
    const zoom = para.getBoundingClientRect().width / para.offsetWidth || 1
    interface Char extends RunBox {
      at: number
      top: number
      bottom: number
      font: string
    }
    const styles = new Map<Element, { font: string; size: number; raised: boolean; ea: boolean }>()
    const chars: Char[] = []
    node.forEach((child, offset) => {
      if (!child.isText || !child.text) return
      for (let i = 0; i < child.text.length; i++) {
        const at = pos + 1 + offset + i
        const { node: dom, offset: domOffset } = view.domAtPos(at, 1)
        if (!(dom instanceof Text) || domOffset >= dom.data.length || !dom.parentElement) continue
        const host = dom.parentElement
        let style = styles.get(host)
        if (!style) {
          const cs = getComputedStyle(host)
          style = {
            font: `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`,
            size: parseFloat(cs.fontSize) || 0,
            raised: cs.verticalAlign !== 'baseline',
            ea: EAST_ASIAN_FACE.test(cs.fontFamily.split(',')[0]),
          }
          styles.set(host, style)
        }
        // superscript and subscript keep their own position
        if (style.raised || style.size === 0) continue
        const range = document.createRange()
        range.setStart(dom, domOffset)
        range.setEnd(dom, domOffset + 1)
        const rect = range.getBoundingClientRect()
        if (rect.height === 0) continue
        // the offset this pass applied last time is taken back out
        let applied = 0
        for (let el: Element | null = host; el && el !== para; el = el.parentElement)
          if (el.classList.contains('doc-ta-shift'))
            applied += parseFloat((el as HTMLElement).style.top) || 0
        const box = this.boxOf(style.font)
        chars.push({
          at,
          top: rect.top / zoom - applied,
          bottom: rect.bottom / zoom - applied,
          font: style.font,
          size: style.size,
          ascent: box.ascent,
          descent: box.descent,
          eastAsian: style.ea,
        })
      }
    })
    const out: Shift[] = []
    let line: Char[] = []
    let lineTop = 0
    let lineBottom = 0
    const flush = () => {
      const big = line.reduce<Char | null>((b, c) => (!b || c.size > b.size ? c : b), null)
      if (!big) return
      for (const c of line) {
        if (c.font === big.font) continue
        const up = Math.round(textAlignmentShift(mode, big, c) * 100) / 100
        if (Math.abs(up) < 0.25) continue
        const last = out[out.length - 1]
        if (last && last.to === c.at && last.up === up) last.to = c.at + 1
        else out.push({ from: c.at, to: c.at + 1, up })
      }
    }
    for (const c of chars) {
      // on one baseline the smaller box lies inside the taller one, whichever comes first
      const overlap = Math.min(c.bottom, lineBottom) - Math.max(c.top, lineTop)
      if (line.length > 0 && overlap > Math.min(c.bottom - c.top, lineBottom - lineTop) / 2) {
        lineTop = Math.min(lineTop, c.top)
        lineBottom = Math.max(lineBottom, c.bottom)
        line.push(c)
        continue
      }
      flush()
      line = [c]
      lineTop = c.top
      lineBottom = c.bottom
    }
    flush()
    return out
  }
}

export const TextAlignmentExtension = Extension.create({
  name: 'textAlignment',
  addProseMirrorPlugins() {
    return [
      new Plugin<DecorationSet>({
        key: textAlignmentKey,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, old) {
            const meta = tr.getMeta(textAlignmentKey) as Decoration[] | undefined
            if (meta)
              return meta.length > 0 ? DecorationSet.create(tr.doc, meta) : DecorationSet.empty
            return old.map(tr.mapping, tr.doc)
          },
        },
        props: {
          decorations(state) {
            return this.getState(state)
          },
        },
        view: (view) => new TextAlignmentView(view),
      }),
    ]
  },
})
