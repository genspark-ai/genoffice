import { AllSelection, Plugin, TextSelection } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'

/**
 * Fill the inter-script seam inside a native text selection.
 *
 * Chromium paints ::selection per selected glyph box, so the gap between a CJK
 * glyph and a Latin/digit glyph stays unpainted. The gap is a collapsible space
 * in `.doc-autospace-pad::before` (word-spacing), so it has no computed width:
 * measure it from the two glyph rects around the pad instead. Only the gap
 * column is painted, with the system colour `Highlight`; the native highlight,
 * paragraph/cell shading, find results, tracked-change tints, dark mode and
 * print are untouched.
 */

export interface SeamBox {
  left: number
  top: number
  width: number
  height: number
}

export interface Rect {
  left: number
  top: number
  width: number
  height: number
}

export interface Bounds {
  left: number
  top: number
  right: number
  bottom: number
}

const PAD_SELECTOR = '.doc-autospace-pad'

/** A text/all selection gets seam fills; object and cell selections do not. */
export const ownsHighlight = (state: EditorView['state']): boolean => {
  const sel = state?.selection
  return !!sel && !sel.empty && (sel instanceof TextSelection || sel instanceof AllSelection)
}

/** The editor scroller's box: seam ticks must never escape it. */
export function seamBounds(view: EditorView): Bounds {
  const scroller = view.dom.closest('.editor-scroll')
  const r = scroller?.getBoundingClientRect()
  if (r && r.width > 0 && r.height > 0) {
    return { left: r.left, top: r.top, right: r.right, bottom: r.bottom }
  }
  return { left: 0, top: 0, right: window.innerWidth, bottom: window.innerHeight }
}

/** First/last text node under an element (inline spans can wrap the glyph). */
function textNodeIn(el: Node, last: boolean): Text | null {
  if (el.nodeType === 3) {
    const t = el as Text
    if (t.length > 0) return t
  }
  let node = last ? el.lastChild : el.firstChild
  while (node) {
    const found = textNodeIn(node, last)
    if (found) return found
    node = last ? node.previousSibling : node.nextSibling
  }
  return null
}

/** Rect of one character; collapsed/zero-size characters have no rect. */
function charRect(node: Node | null, offset: number): Rect | null {
  if (!node) return null
  const len = node.nodeType === 3 ? (node as Text).length : 1
  if (len === 0) return null
  const at = Math.min(Math.max(offset, 0), len - 1)
  const range = document.createRange()
  try {
    range.setStart(node, at)
    range.setEnd(node, at + 1)
  } catch {
    return null
  }
  for (const r of range.getClientRects?.() ?? [])
    if (r.height > 0) return { left: r.left, top: r.top, width: r.width, height: r.height }
  return null
}

/** Last glyph before the pad: its right edge is where the gap starts. */
function prevCharRect(pad: HTMLElement): Rect | null {
  let node: Node | null = pad.previousSibling
  while (node) {
    if (node.nodeType === 3) {
      const t = node as Text
      if (t.length > 0) return charRect(t, t.length - 1)
    } else if (node.nodeType === 1) {
      const t = textNodeIn(node, true)
      if (t) return charRect(t, t.length - 1)
    }
    node = node.previousSibling
  }
  return null
}

/** First glyph inside the pad: its left edge is where the gap ends. */
function padCharRect(pad: HTMLElement): Rect | null {
  const t = textNodeIn(pad, false)
  return t ? charRect(t, 0) : null
}

/** A gap wider than this is not an autospace seam (a wrapped space, a tab). */
const MAX_SEAM_PX = 24

/**
 * Rectangles covering the gap column of every pad the range touches. Pads are
 * looked up only inside blocks that intersect the scroller, so a select-all on
 * a long document does not measure the whole pad population.
 *
 * `posAtDom`/`ranges` are injectable so jsdom tests can supply geometry.
 */
export function collectSeamRects(
  root: ParentNode & Node,
  ranges: readonly Range[],
  bounds: Bounds,
  selectionFrom: number,
  posAtDom: (pad: HTMLElement) => number,
  rectOf: (el: HTMLElement) => Rect = (el) => el.getBoundingClientRect(),
): SeamBox[] {
  if (ranges.length === 0) return []
  const out: SeamBox[] = []
  const blocks =
    root.nodeType === 1
      ? (Array.from((root as Element).children) as HTMLElement[])
      : (Array.from(root.querySelectorAll?.('*') ?? []) as HTMLElement[]).filter(
          (el) => el.parentElement === root,
        )

  for (const block of blocks) {
    const br = rectOf(block)
    if (br.width <= 0 || br.height <= 0) continue
    if (br.top >= bounds.bottom || br.top + br.height <= bounds.top) continue
    for (const pad of Array.from(block.querySelectorAll<HTMLElement>(PAD_SELECTOR))) {
      const range = ranges.find((candidate) => candidate.intersectsNode(pad))
      if (!range) continue
      // A pad that starts the selection has its gap before the selected text;
      // the native highlight already covers the first glyph.
      if (posAtDom(pad) >= selectionFrom) continue
      const before = prevCharRect(pad)
      const after = padCharRect(pad)
      if (!before || !after) continue
      const width = after.left - (before.left + before.width)
      if (!(width > 0) || width > MAX_SEAM_PX) continue
      const r = rectOf(pad)
      if (r.width <= 0 || r.height <= 0) continue
      const left = before.left + before.width
      if (left >= bounds.right || left + width <= bounds.left) continue
      out.push({ left, top: after.top, width, height: after.height })
    }
  }
  return out
}

export function selectionHighlightOverlay(): Plugin {
  return new Plugin({
    view(view: EditorView) {
      const layer = document.createElement('div')
      layer.className = 'go-selection-seams'
      document.body.appendChild(layer)

      const selectionRanges = (): Range[] => {
        const ranges: Range[] = []
        for (const range of view.state.selection.ranges) {
          const start = view.domAtPos(range.$from.pos)
          const end = view.domAtPos(range.$to.pos)
          const domRange = document.createRange()
          try {
            domRange.setStart(start.node, start.offset)
            domRange.setEnd(end.node, end.offset)
          } catch {
            continue // positions can be transiently unmappable mid-transaction
          }
          ranges.push(domRange)
        }
        return ranges
      }

      let frame = 0
      const build = () => {
        frame = 0
        layer.replaceChildren()
        if (!view.dom.isConnected) return
        // Only paint while the real DOM selection is inside the editor: a focus
        // move to a ribbon input or a sub-editor clears the native highlight,
        // and ticks left behind would imply a selection that is not there.
        const active = document.hasFocus() && view.dom.classList.contains('ProseMirror-focused')
        if (!active || !ownsHighlight(view.state)) return
        const rects = collectSeamRects(
          view.dom,
          selectionRanges(),
          seamBounds(view),
          view.state.selection.from,
          (pad) => view.posAtDOM(pad, 0),
        )
        for (const r of rects) {
          const el = document.createElement('div')
          el.className = 'go-selection-seam'
          el.style.left = `${r.left}px`
          el.style.top = `${r.top}px`
          el.style.width = `${r.width}px`
          el.style.height = `${r.height}px`
          layer.appendChild(el)
        }
      }

      const schedule = () => {
        if (!frame) frame = requestAnimationFrame(build)
      }
      const onScroll = () => schedule()
      document.addEventListener('selectionchange', schedule)
      window.addEventListener('scroll', onScroll, true)
      window.addEventListener('resize', schedule)
      window.addEventListener('focus', schedule)
      window.addEventListener('blur', schedule)
      document.addEventListener('focusin', schedule)
      document.addEventListener('focusout', schedule)
      schedule()

      return {
        update: schedule,
        destroy() {
          if (frame) cancelAnimationFrame(frame)
          document.removeEventListener('selectionchange', schedule)
          window.removeEventListener('scroll', onScroll, true)
          window.removeEventListener('resize', schedule)
          window.removeEventListener('focus', schedule)
          window.removeEventListener('blur', schedule)
          document.removeEventListener('focusin', schedule)
          document.removeEventListener('focusout', schedule)
          layer.remove()
        },
      }
    },
  })
}
