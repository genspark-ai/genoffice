import { Extension } from '@tiptap/core'
import { Plugin, PluginKey, Selection, TextSelection } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'

const key = new PluginKey('pageGapArrowNav')

/** nearest scrollable ancestor (the docs canvas scroller), viewport as fallback */
function scrollableAncestor(el: HTMLElement | null): HTMLElement | null {
  for (let n = el; n; n = n.parentElement) {
    if (n.scrollHeight > n.clientHeight + 1) {
      const oy = getComputedStyle(n).overflowY
      if (oy === 'auto' || oy === 'scroll') return n
    }
  }
  return (document.scrollingElement as HTMLElement | null) ?? null
}

type Rect = { left: number; right: number; top: number; bottom: number }

/**
 * Where the caret is actually rendered. At a line-wrap boundary the head
 * position is ambiguous between two lines, and coordsAtPos's side bias can
 * pick the wrong one — native vertical motion moves from the line the BROWSER
 * renders the caret on (r122: after Shift+Right + Shift+Downs the head sat on
 * the wrap boundary just above the gap; the estimate said "not adjacent", the
 * plugin declined, and the native move blew the selection up to the document
 * end). The live DOM selection's focus rect resolves the ambiguity exactly;
 * PM's estimate is the fallback when the DOM selection is absent or elsewhere.
 */
function caretRect(view: EditorView, head: number, dir: -1 | 1): Rect {
  const domSel = (view.root as Document).getSelection?.()
  if (domSel?.focusNode && view.dom.contains(domSel.focusNode)) {
    try {
      const range = document.createRange()
      range.setStart(domSel.focusNode, domSel.focusOffset)
      range.collapse(true)
      const r = range.getClientRects()[0] ?? range.getBoundingClientRect()
      if (r && r.height > 0) return r
    } catch {
      /* element-boundary offsets: fall through to the estimate */
    }
  }
  return view.coordsAtPos(head, dir === 1 ? -1 : 1)
}

/**
 * Vertical caret motion across page gaps (alpha ledger r118/r122): the
 * inter-page gap is a large contentEditable=false widget, and Chromium's
 * native ArrowDown/ArrowUp — including their Shift-extension — give up on it
 * when the page break falls MID-paragraph (inline gap inside the textblock):
 * the caret or selection head jumps to the end/start of the whole document
 * instead of the adjacent line on the next/previous page. When the gap is the
 * next visual line in the pressed direction, place the caret at the same x
 * just past it ourselves; everywhere else the browser's native motion (with
 * its goal-column memory) stays in charge.
 */
function crossPageGap(view: EditorView, dir: -1 | 1, extend: boolean): boolean {
  const { selection, doc } = view.state
  if (!(selection instanceof TextSelection)) return false
  const caret = caretRect(view, selection.head, dir)
  const lineH = Math.max(caret.bottom - caret.top, 8)
  const x = (caret.left + caret.right) / 2
  // nearest gap in the pressed direction at this x (rect scan, not
  // elementFromPoint: the gap may sit outside the viewport)
  let gap: HTMLElement | null = null
  let gapRect: DOMRect | null = null
  for (const el of Array.from(view.dom.querySelectorAll<HTMLElement>('.page-gap'))) {
    const r = el.getBoundingClientRect()
    if (r.height <= 0 || x < r.left - 2 || x > r.right + 2) continue
    if (dir === 1 ? r.top < caret.top : r.bottom > caret.bottom) continue
    if (!gapRect || (dir === 1 ? r.top < gapRect.top : r.bottom > gapRect.bottom)) {
      gap = el
      gapRect = r
    }
  }
  if (!gap || !gapRect) return false
  // posAtCoords is viewport hit-testing and yields null/degenerate positions
  // for points outside the scrollport (same trap App's posFromAnchor
  // documents) — EVERY hit-tested point below must first be scrolled into
  // view. Decline paths restore the scroll; a successful dispatch re-scrolls
  // via scrollIntoView anyway.
  const scroller = scrollableAncestor(view.dom)
  const savedScrollTop = scroller?.scrollTop ?? 0
  const restore = () => {
    if (scroller) scroller.scrollTop = savedScrollTop
  }
  const ensureVisible = (y: number): number => {
    if (!scroller) return 0
    const s = scroller.getBoundingClientRect()
    const top = Math.max(s.top, 0)
    const bottom = Math.min(s.bottom, window.innerHeight)
    if (y >= top + lineH && y <= bottom - lineH) return 0
    const before = scroller.scrollTop
    scroller.scrollTop += y - (top + bottom) / 2
    return scroller.scrollTop - before
  }
  // a real text line between the caret line and the gap means native handles
  // this press. Tested by hit-testing the band's midpoint rather than probing
  // a fixed depth: the gap's own line box carries leading, so the gap can sit
  // a full line-height below the last text line with nothing in between (r122)
  let caretTop = caret.top
  let caretBottom = caret.bottom
  const band = () =>
    dir === 1
      ? { top: caretBottom, bottom: gapRect!.top }
      : { top: gapRect!.bottom, bottom: caretTop }
  let b = band()
  if (b.bottom - b.top > 2) {
    if (ensureVisible((b.top + b.bottom) / 2) !== 0) {
      // rects moved with the scroll: re-read both ends of the band
      const c = caretRect(view, selection.head, dir)
      caretTop = c.top
      caretBottom = c.bottom
      gapRect = gap.getBoundingClientRect()
      b = band()
    }
    const mid = view.posAtCoords({ left: x, top: (b.top + b.bottom) / 2 })
    if (!mid) {
      // still unresolvable: declining can at worst leave native to move one
      // line; jumping could skip every remaining line on the page (bugbot)
      restore()
      return false
    }
    const r = view.coordsAtPos(mid.pos)
    if (r.top > b.top - 2 && r.bottom < b.bottom + 2) {
      restore()
      return false
    }
  }
  const landingY = (r: DOMRect) => (dir === 1 ? r.bottom + lineH * 0.5 : r.top - lineH * 0.5)
  let targetY = landingY(gapRect)
  if (ensureVisible(targetY) !== 0) targetY = landingY(gap.getBoundingClientRect())
  const target = view.posAtCoords({ left: x, top: targetY })
  if (!target) {
    restore()
    return false
  }
  const next = extend
    ? TextSelection.create(doc, selection.anchor, target.pos)
    : Selection.near(doc.resolve(target.pos), dir)
  if (next.head === selection.head) {
    restore()
    return false
  }
  view.dispatch(view.state.tr.setSelection(next).scrollIntoView())
  return true
}

export const PageGapNavExtension = Extension.create({
  name: 'pageGapArrowNav',

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key,
        props: {
          handleKeyDown(view, event) {
            if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return false
            if (event.altKey || event.ctrlKey || event.metaKey) return false
            return crossPageGap(view, event.key === 'ArrowDown' ? 1 : -1, event.shiftKey)
          },
        },
      }),
    ]
  },
})
