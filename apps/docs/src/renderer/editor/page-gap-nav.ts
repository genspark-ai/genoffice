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
  const estimate = view.coordsAtPos(head, dir === 1 ? -1 : 1)
  const domSel = (view.root as Document).getSelection?.()
  if (domSel?.focusNode && view.dom.contains(domSel.focusNode)) {
    // A focus parked ON the gap widget (a click near the page seam leaves it
    // there) yields the gap's own multi-hundred-px box as the "caret" rect;
    // every downstream computation then runs on garbage and the dispatched
    // jump can land paragraphs — or the whole document — past the boundary
    // (alpha ledger r143). Only a text-line-sized rect off the gap counts.
    const focusEl =
      domSel.focusNode.nodeType === Node.TEXT_NODE
        ? domSel.focusNode.parentElement
        : (domSel.focusNode as HTMLElement)
    if (!focusEl?.closest?.('.page-gap')) {
      try {
        const range = document.createRange()
        range.setStart(domSel.focusNode, domSel.focusOffset)
        range.collapse(true)
        const r = range.getClientRects()[0] ?? range.getBoundingClientRect()
        const estimateH = Math.max(estimate.bottom - estimate.top, 8)
        if (r && r.height > 0 && r.height <= estimateH * 3) return r
      } catch {
        /* element-boundary offsets: fall through to the estimate */
      }
    }
  }
  return estimate
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
/** consume escape valve (r147): absorbing a press twice for the SAME
 *  selection head means the geometry error is persistent, not a transient
 *  reflow race — hand the press back to native rather than freezing the key.
 *  One absorbed tick covers the r143 race; the next press either recomputes
 *  cleanly (race) or goes native (persistent). */
let lastConsumedHead = -1
const consumePress = (head: number): boolean => {
  const firstTime = head !== lastConsumedHead
  lastConsumedHead = head
  return firstTime
}

function crossPageGap(view: EditorView, dir: -1 | 1, extend: boolean): boolean {
  const { selection, doc } = view.state
  if (!(selection instanceof TextSelection)) return false
  const caret = caretRect(view, selection.head, dir)
  const lineH = Math.max(caret.bottom - caret.top, 8)
  // Horizontal comes from the POSITION estimate, not the DOM rect: under key
  // repeat the DOM selection rect can be read mid-reflow and report a stale x
  // at the page margin — every hit-test below then probes the wrong column,
  // the band test passes vacuously and the landing resolves garbage (alpha
  // ledger r143 reopen: page 2 selections jumped to page 4). The DOM rect
  // stays authoritative for the VERTICAL line only (the r122 wrap-boundary
  // ambiguity it was introduced for).
  const estimate = view.coordsAtPos(selection.head, dir === 1 ? -1 : 1)
  const x = (estimate.left + estimate.right) / 2
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
  let landingGapRect: DOMRect = gapRect
  let targetY = landingY(landingGapRect)
  if (ensureVisible(targetY) !== 0) {
    landingGapRect = gap.getBoundingClientRect()
    targetY = landingY(landingGapRect)
  }
  const target = view.posAtCoords({ left: x, top: targetY })
  if (!target) {
    // the gap IS the adjacent line (the band check passed) — native cannot
    // cross an inline gap and would jump a whole page, so CONSUME the press:
    // one skipped key-repeat tick, and the next press recomputes fresh
    restore()
    return consumePress(selection.head)
  }
  // the landing must be the line adjacent to the gap: a hit-test that
  // resolves further away (stale geometry, margins) would dispatch a
  // multi-paragraph jump (r143)
  const landed = view.coordsAtPos(target.pos)
  if (
    dir === 1
      ? landed.top > landingGapRect.bottom + lineH * 2
      : landed.bottom < landingGapRect.top - lineH * 2
  ) {
    restore()
    return consumePress(selection.head)
  }
  const next = extend
    ? TextSelection.create(doc, selection.anchor, target.pos)
    : Selection.near(doc.resolve(target.pos), dir)
  if (next.head === selection.head) {
    restore()
    return false
  }
  // a successful crossing clears the absorb marker: a LATER transient
  // failure at this same head must get its skipped tick again instead of
  // falling straight to native (bugbot)
  lastConsumedHead = -1
  view.dispatch(view.state.tr.setSelection(next).scrollIntoView())
  return true
}

/** A vertical arrow press handed to native, with the gap doc-positions at
 *  press time: whatever native does, ONE press may cross at most ONE page
 *  boundary. Rect-based guards keep failing in ways we cannot fully
 *  enumerate (alpha ledger r143, twice reopened) — document positions are
 *  the only stable ground truth, so the invariant is enforced after the
 *  fact on native's own transaction. */
let nativeArrow: {
  head: number
  dir: -1 | 1
  extend: boolean
  anchor: number
  gapPositions: number[]
  at: number
} | null = null

function recordNativeArrow(view: EditorView, dir: -1 | 1, extend: boolean): void {
  const { selection } = view.state
  if (!(selection instanceof TextSelection)) {
    nativeArrow = null
    return
  }
  const gapPositions: number[] = []
  for (const el of Array.from(view.dom.querySelectorAll<HTMLElement>('.page-gap'))) {
    try {
      gapPositions.push(view.posAtDOM(el, 0))
    } catch {
      /* detached widget */
    }
  }
  gapPositions.sort((a, b) => a - b)
  nativeArrow = {
    head: selection.head,
    dir,
    extend,
    anchor: selection.anchor,
    gapPositions,
    at: Date.now(),
  }
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
            const dir = event.key === 'ArrowDown' ? 1 : -1
            const handled = crossPageGap(view, dir, event.shiftKey)
            if (handled) nativeArrow = null
            else recordNativeArrow(view, dir, event.shiftKey)
            return handled
          },
        },
        appendTransaction(transactions, oldState, newState) {
          const pending = nativeArrow
          if (!pending) return null
          if (!transactions.some((tr) => tr.selectionSet)) return null
          // only the native motion right after the recorded press qualifies
          if (Date.now() - pending.at > 300) {
            nativeArrow = null
            return null
          }
          nativeArrow = null
          const selection = newState.selection
          if (!(selection instanceof TextSelection)) return null
          if (oldState.selection.head !== pending.head) return null
          const moved = selection.head - pending.head
          if (pending.dir === 1 ? moved <= 0 : moved >= 0) return null
          // gaps between the old and new head. Inline gaps sit AT the first
          // position of the new page, so the endpoint on the far side of the
          // motion counts too: landing exactly on a downward gap position (or
          // starting exactly on one going up) is a real crossing (bugbot)
          const crossed = pending.gapPositions.filter((gapPos) =>
            pending.dir === 1
              ? gapPos > pending.head && gapPos <= selection.head
              : gapPos <= pending.head && gapPos > selection.head,
          )
          if (crossed.length <= 1) return null
          // eslint-disable-next-line no-console -- deliberate breadcrumb: the
          // clamp firing means a native jump crossed 2+ page boundaries; when
          // a user reports a jump we can ask for this line from their console
          console.warn('[gap-nav] clamped a multi-boundary native jump', {
            from: pending.head,
            to: selection.head,
            crossed: crossed.length,
          })
          // clamp to just past the FIRST crossed boundary. The gap position
          // itself allows a cursor (inline widget), where Selection.near's
          // bias is a no-op — nudge one position in the motion direction so
          // an upward clamp really lands on the previous page's last line
          // and a downward one on the new page's first (bugbot)
          const firstGap = pending.dir === 1 ? Math.min(...crossed) : Math.max(...crossed)
          // never nudge onto or past the NEXT boundary: a degenerate one-
          // position page would otherwise still be skipped (bugbot) — in that
          // case stay on the gap position itself
          const neighborGap =
            pending.dir === 1
              ? Math.min(...pending.gapPositions.filter((gapPos) => gapPos > firstGap), Infinity)
              : Math.max(...pending.gapPositions.filter((gapPos) => gapPos < firstGap), -Infinity)
          const nudgedCandidate =
            pending.dir === 1
              ? Math.min(newState.doc.content.size, firstGap + 1)
              : Math.max(0, firstGap - 1)
          // Asymmetric bounds: a gap anchors at the FIRST position of the
          // page below it. Downward, reaching the next gap means we entered
          // the page after next — overshoot at equality. Upward, landing ON
          // the previous gap position IS the destination (it's that page's
          // first position), so only going strictly below it overshoots
          // (bugbot: one-position pages were skipped on upward clamps).
          const nudged =
            pending.dir === 1
              ? nudgedCandidate >= neighborGap
                ? firstGap
                : nudgedCandidate
              : nudgedCandidate < neighborGap
                ? firstGap
                : nudgedCandidate
          const clampedHead = Selection.near(newState.doc.resolve(nudged), pending.dir).head
          const next = pending.extend
            ? TextSelection.create(newState.doc, pending.anchor, clampedHead)
            : Selection.near(newState.doc.resolve(clampedHead), pending.dir)
          return newState.tr.setSelection(next).scrollIntoView()
        },
      }),
    ]
  },
})
