/**
 * Floating header pictures (the page-hf-float-img layer) moved and resized
 * like Word: in header editing a press selects one (frame and handles) and a
 * drag moves it; a double-click on one from the body opens header editing
 * with it selected. The layer stays pointer-events: none (body paragraphs
 * paint over it), so presses are hit-tested against the pictures' boxes.
 */
import type { HfPictureEdit } from '@genoffice/docx-engine'
import type { HfView } from '../doc-state'

export interface HfFloatTarget {
  section: number
  variant: HfView
  pageIdx: number
  drawing: number
}

/** a picture's box in the strip: page-relative top-left plus size, CSS px */
type HfBox = Required<HfPictureEdit>

type Handle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w'
const HANDLES: readonly Handle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']
const DRAG_THRESHOLD_PX = 3
const MIN_SIDE_PX = 12

/** what a float element carries for editing (set by tagHfFloatEl / makeHfFloatImgEl) */
function targetOf(el: HTMLElement): HfFloatTarget | null {
  const { hfDrawing, hfSection, hfVariant, hfPageIdx } = el.dataset
  if (hfDrawing == null || hfSection == null || hfPageIdx == null) return null
  return {
    section: Number(hfSection),
    variant: (hfVariant as HfView | undefined) ?? 'default',
    pageIdx: Number(hfPageIdx),
    drawing: Number(hfDrawing),
  }
}

/** a float layer element as a header strip of `section` / `variant` on page `pageIdx` shows it */
export function tagHfFloatEl(
  el: HTMLElement,
  strip: { section: number; variant: HfView },
  pageIdx: number,
): HTMLElement {
  el.dataset.hfSection = String(strip.section)
  el.dataset.hfVariant = strip.variant
  el.dataset.hfPageIdx = String(pageIdx)
  return el
}

/** CSS zoom of the page canvas around el: client px divided by it are CSS px */
function zoomOf(el: Element): number {
  const zoomEl = el.closest('.doc-zoom') as HTMLElement | null
  return zoomEl ? parseFloat(getComputedStyle(zoomEl).zoom || '1') || 1 : 1
}

/** the box the element is laid out at (makeHfFloatImgEl) */
function boxOf(el: HTMLElement): HfBox | null {
  const x = Number(el.dataset.pageX)
  const y = Number(el.dataset.pageY)
  const w = Number(el.dataset.boxW) || el.offsetWidth
  const h = Number(el.dataset.boxH) || el.offsetHeight
  if (!Number.isFinite(x) || !Number.isFinite(y) || !(w > 0) || !(h > 0)) return null
  return { posXPx: x, posYPx: y, widthPx: w, heightPx: h }
}

const round1 = (v: number) => Math.round(v * 10) / 10

export function installHfFloatEdit(opts: {
  /** the scroller holding the pages */
  root: HTMLElement
  /** a header/footer strip editor is open */
  editing: () => boolean
  /** the document can be edited (not read-only / protected) */
  editable: () => boolean
  /** open header editing on page `pageIdx` */
  open: (pageIdx: number) => void
  /** false when the edit could not be taken (the strip has no target) */
  commit: (target: HfFloatTarget, edit: HfPictureEdit) => boolean
}): () => void {
  const { root } = opts
  let selected: HfFloatTarget | null = null
  /** a double-click from the body asked for header editing: the selection waits for it */
  let awaitingOpen = false
  let frame: HTMLDivElement | null = null
  let raf = 0
  let swallowDblClick = false
  /** a committed edit drawn on its element until pagination replaces the element */
  let shown: {
    el: HTMLElement
    target: HfFloatTarget
    box: HfBox
    restore: () => void
  } | null = null

  const floats = () => Array.from(root.querySelectorAll<HTMLElement>('.page-hf-float-img'))
  /** the picture under the point; in-front pictures beat behind-text ones, later ones earlier ones */
  const hitAt = (x: number, y: number, includeBehind: boolean): HTMLElement | null => {
    let best: HTMLElement | null = null
    let bestFront = false
    for (const el of floats()) {
      if (!targetOf(el)) continue
      const front = el.dataset.hfBehind == null
      if (!front && !includeBehind) continue
      const r = el.getBoundingClientRect()
      if (x < r.left || x > r.right || y < r.top || y > r.bottom) continue
      if (!best || front || !bestFront) {
        best = el
        bestFront = front
      }
    }
    return best
  }
  const elOf = (t: HfFloatTarget) =>
    root.querySelector<HTMLElement>(
      `.page-hf-float-img[data-hf-page-idx="${t.pageIdx}"][data-hf-drawing="${t.drawing}"]`,
    )

  // the frame follows its picture through scrolling, zoom, layout and the
  // pagination pass that rebuilds the layer, re-placed once per frame at most
  const schedule = () => {
    if (!raf) raf = requestAnimationFrame(place)
  }
  let mutations: MutationObserver | null = null
  let resizes: ResizeObserver | null = null
  const watch = () => {
    const want = selected != null || shown != null
    if (want === (mutations != null)) return
    if (want) {
      // class: header editing opening / closing; style: zoom, previews, re-placed layers
      mutations = new MutationObserver(schedule)
      mutations.observe(root, {
        subtree: true,
        childList: true,
        attributes: true,
        attributeFilter: ['class', 'style'],
      })
      resizes = typeof ResizeObserver === 'undefined' ? null : new ResizeObserver(schedule)
      resizes?.observe(root)
      root.addEventListener('scroll', schedule, true)
      window.addEventListener('resize', schedule)
    } else {
      mutations?.disconnect()
      resizes?.disconnect()
      mutations = resizes = null
      root.removeEventListener('scroll', schedule, true)
      window.removeEventListener('resize', schedule)
    }
  }

  const frameEl = (): HTMLDivElement => {
    if (frame) return frame
    frame = document.createElement('div')
    frame.className = 'doc-run-pic-frame hf-float-frame'
    for (const dir of HANDLES) {
      const h = document.createElement('span')
      h.className = `img-resize-handle img-resize-handle-${dir}`
      h.dataset.resizeHandle = dir
      h.setAttribute('aria-hidden', 'true')
      frame.append(h)
    }
    frame.addEventListener('mousedown', onHandleDown)
    document.body.append(frame)
    return frame
  }
  const place = () => {
    raf = 0
    // pagination swapped in the re-placed element: the old one took the preview along
    if (shown && (!shown.el.isConnected || elOf(shown.target) !== shown.el)) {
      if (shown.el.isConnected) shown.restore()
      shown = null
    }
    if (selected) {
      if (opts.editing()) awaitingOpen = false
      else if (!awaitingOpen) selected = null
    }
    watch()
    const el = selected && !awaitingOpen ? elOf(selected) : null
    const rect = el?.getBoundingClientRect()
    if (!rect || !(rect.width > 0) || !(rect.height > 0)) {
      if (frame) frame.hidden = true
      return
    }
    const box = frameEl()
    box.hidden = false
    box.style.left = `${rect.left}px`
    box.style.top = `${rect.top}px`
    box.style.width = `${rect.width}px`
    box.style.height = `${rect.height}px`
    // fixed over the window: handles scrolled under the ribbon / status bar stay hidden
    const port = root.getBoundingClientRect()
    for (const handle of Array.from(box.children) as HTMLElement[]) {
      const r = handle.getBoundingClientRect()
      const cx = r.left + r.width / 2
      const cy = r.top + r.height / 2
      const visible = cx >= port.left && cx <= port.right && cy >= port.top && cy <= port.bottom
      handle.style.visibility = visible ? '' : 'hidden'
    }
  }
  const select = (t: HfFloatTarget | null) => {
    selected = t
    if (!t) awaitingOpen = false
    watch()
    schedule()
  }

  /** drag (move or resize) with a live preview on the element itself */
  const track = (
    e: MouseEvent,
    el: HTMLElement,
    target: HfFloatTarget,
    next: (start: HfBox, dx: number, dy: number) => HfBox,
    threshold: number,
  ) => {
    // where the element is laid out, and where the picture shows: an edit
    // committed a moment ago is still a preview until the pass re-places it
    const laid = boxOf(el)
    if (!laid) return
    const pending = shown?.el === el ? shown : null
    const start = pending?.box ?? laid
    const zoom = zoomOf(el)
    const x0 = e.clientX
    const y0 = e.clientY
    let active = threshold === 0
    const prior = { translate: el.style.translate, width: el.style.width, height: el.style.height }
    const restore =
      pending?.restore ??
      (() => {
        el.style.translate = prior.translate
        el.style.width = prior.width
        el.style.height = prior.height
      })
    if (pending) shown = null
    const draw = (box: HfBox) => {
      el.style.width = `${box.widthPx}px`
      el.style.height = `${box.heightPx}px`
      el.style.translate = `${box.posXPx - laid.posXPx}px ${box.posYPx - laid.posYPx}px`
    }
    const editAt = (ev: MouseEvent) =>
      next(start, (ev.clientX - x0) / zoom, (ev.clientY - y0) / zoom)
    const onMove = (ev: MouseEvent) => {
      if (!active) {
        if (Math.hypot(ev.clientX - x0, ev.clientY - y0) < threshold) return
        active = true
      }
      document.body.classList.add('docs-run-pic-dragging')
      draw(editAt(ev))
    }
    const onUp = (ev: MouseEvent) => {
      window.removeEventListener('mousemove', onMove, true)
      window.removeEventListener('mouseup', onUp, true)
      document.body.classList.remove('docs-run-pic-dragging')
      const box = active ? editAt(ev) : null
      const movedX = box != null && Math.abs(box.posXPx - start.posXPx) >= 0.5
      const movedY = box != null && Math.abs(box.posYPx - start.posYPx) >= 0.5
      const resized =
        box != null &&
        (Math.abs(box.widthPx - start.widthPx) >= 0.5 ||
          Math.abs(box.heightPx - start.heightPx) >= 0.5)
      if (!box || !(movedX || movedY || resized)) {
        if (pending) shown = pending
        else restore()
        return
      }
      // only the axes the corner moved along: a resize from the far edge
      // leaves the picture's own position (base and offset) alone
      const edit: HfPictureEdit = {
        widthPx: Math.max(1, Math.round(box.widthPx)),
        heightPx: Math.max(1, Math.round(box.heightPx)),
      }
      if (movedX) edit.posXPx = round1(box.posXPx)
      if (movedY) edit.posYPx = round1(box.posYPx)
      if (!opts.commit(target, edit)) {
        restore()
        return
      }
      // the preview stays until pagination swaps in the re-placed element
      shown = { el, target, box, restore }
      watch()
    }
    window.addEventListener('mousemove', onMove, true)
    window.addEventListener('mouseup', onUp, true)
  }

  const startMove = (e: MouseEvent, el: HTMLElement, target: HfFloatTarget) =>
    track(
      e,
      el,
      target,
      (s, dx, dy) => ({ ...s, posXPx: s.posXPx + dx, posYPx: s.posYPx + dy }),
      DRAG_THRESHOLD_PX,
    )

  function onHandleDown(e: MouseEvent) {
    if (e.button !== 0 || !selected) return
    const dir = (e.target as HTMLElement).closest<HTMLElement>('.img-resize-handle')?.dataset
      .resizeHandle as Handle | undefined
    const el = elOf(selected)
    if (!dir || !el) return
    // keeps the focus (and so header editing) in the strip editor
    e.preventDefault()
    e.stopPropagation()
    const west = dir.includes('w')
    const east = dir.includes('e')
    const north = dir.includes('n')
    const south = dir.includes('s')
    track(
      e,
      el,
      selected,
      (s, dx, dy) => {
        let w = s.widthPx
        let h = s.heightPx
        if ((west || east) && (north || south)) {
          // corners keep the aspect ratio (Word)
          const delta =
            (dx * (west ? -1 : 1) * w + dy * (north ? -1 : 1) * h) / Math.max(1, w * w + h * h)
          const scale = Math.max(MIN_SIDE_PX / w, MIN_SIDE_PX / h, 1 + delta)
          w *= scale
          h *= scale
        } else if (west || east) w = Math.max(MIN_SIDE_PX, w + (west ? -dx : dx))
        else h = Math.max(MIN_SIDE_PX, h + (north ? -dy : dy))
        // the opposite edge stays put
        return {
          posXPx: west ? s.posXPx + s.widthPx - w : s.posXPx,
          posYPx: north ? s.posYPx + s.heightPx - h : s.posYPx,
          widthPx: w,
          heightPx: h,
        }
      },
      0,
    )
  }

  const onMouseDown = (e: MouseEvent) => {
    if (e.button !== 0 || !opts.editable()) return
    const editing = opts.editing()
    // from the body only a double-click reaches a header picture, and only one
    // drawn in front of the text: a behind-text one leaves the word under it to the body
    if (!editing && e.detail < 2) {
      // a click elsewhere drops a selection still waiting for header editing
      if (awaitingOpen) select(null)
      return
    }
    const el = hitAt(e.clientX, e.clientY, editing)
    const target = el ? targetOf(el) : null
    if (!el || !target) {
      if (editing) select(null)
      return
    }
    // keeps the focus where it is: in header editing the strip editor stays open
    e.preventDefault()
    e.stopPropagation()
    if (!editing) {
      swallowDblClick = true
      awaitingOpen = true
      select(target)
      opts.open(target.pageIdx)
      return
    }
    select(target)
    startMove(e, el, target)
  }
  const onDblClick = (e: MouseEvent) => {
    if (!swallowDblClick) return
    swallowDblClick = false
    e.preventDefault()
    e.stopPropagation()
  }
  const onHover = (e: MouseEvent) => {
    if (e.buttons) return
    const over = opts.editing() && opts.editable() && hitAt(e.clientX, e.clientY, true) != null
    root.classList.toggle('hf-float-hover', over)
  }
  const onKey = (e: KeyboardEvent) => {
    if (e.key === 'Escape' && selected) select(null)
  }

  root.addEventListener('mousedown', onMouseDown, true)
  root.addEventListener('dblclick', onDblClick, true)
  root.addEventListener('mousemove', onHover)
  window.addEventListener('keydown', onKey, true)
  return () => {
    cancelAnimationFrame(raf)
    selected = null
    shown = null
    watch()
    root.removeEventListener('mousedown', onMouseDown, true)
    root.removeEventListener('dblclick', onDblClick, true)
    root.removeEventListener('mousemove', onHover)
    window.removeEventListener('keydown', onKey, true)
    root.classList.remove('hf-float-hover')
    frame?.remove()
  }
}
