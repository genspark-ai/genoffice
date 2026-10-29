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

/** the box an edit starts from: page-relative top-left plus size, CSS px */
function boxOf(el: HTMLElement): HfPictureEdit | null {
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
  commit: (target: HfFloatTarget, edit: HfPictureEdit) => void
}): () => void {
  const { root } = opts
  let selected: HfFloatTarget | null = null
  let frame: HTMLDivElement | null = null
  let raf = 0
  let swallowDblClick = false
  /** when a double-click asked for header editing: the selection waits for it to open */
  let openedAt = 0

  const floats = () => Array.from(root.querySelectorAll<HTMLElement>('.page-hf-float-img'))
  /** the picture under the point; in-front pictures beat behind-text ones, later ones earlier ones */
  const hitAt = (x: number, y: number, includeBehind: boolean): HTMLElement | null => {
    let best: HTMLElement | null = null
    let bestFront = false
    for (const el of floats()) {
      if (!targetOf(el)) continue
      const front = el.style.zIndex === '1'
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
    if (selected && !opts.editing() && performance.now() - openedAt > 1000) selected = null
    const el = selected ? elOf(selected) : null
    const rect = el?.getBoundingClientRect()
    if (!rect || !(rect.width > 0) || !(rect.height > 0)) {
      if (frame) frame.hidden = true
      if (selected) raf = requestAnimationFrame(place)
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
      const shown = cx >= port.left && cx <= port.right && cy >= port.top && cy <= port.bottom
      handle.style.visibility = shown ? '' : 'hidden'
    }
    // pagination rebuilds the layer after a commit: keep following it
    raf = requestAnimationFrame(place)
  }
  const select = (t: HfFloatTarget | null) => {
    selected = t
    if (!raf) raf = requestAnimationFrame(place)
  }

  /** drag (move or resize) with a live preview on the element itself */
  const track = (
    e: MouseEvent,
    el: HTMLElement,
    target: HfFloatTarget,
    next: (dx: number, dy: number) => HfPictureEdit | null,
    preview: (edit: HfPictureEdit, start: HfPictureEdit) => void,
    threshold: number,
  ) => {
    const start = boxOf(el)
    if (!start) return
    const zoom = zoomOf(el)
    const x0 = e.clientX
    const y0 = e.clientY
    let active = threshold === 0
    const prior = {
      translate: el.style.translate,
      width: el.style.width,
      height: el.style.height,
    }
    const restore = () => {
      el.style.translate = prior.translate
      el.style.width = prior.width
      el.style.height = prior.height
      document.body.classList.remove('docs-run-pic-dragging')
    }
    const editAt = (ev: MouseEvent) => next((ev.clientX - x0) / zoom, (ev.clientY - y0) / zoom)
    const onMove = (ev: MouseEvent) => {
      if (!active) {
        if (Math.hypot(ev.clientX - x0, ev.clientY - y0) < threshold) return
        active = true
      }
      document.body.classList.add('docs-run-pic-dragging')
      const edit = editAt(ev)
      if (edit) preview(edit, start)
    }
    const onUp = (ev: MouseEvent) => {
      window.removeEventListener('mousemove', onMove, true)
      window.removeEventListener('mouseup', onUp, true)
      const edit = active ? editAt(ev) : null
      const moved =
        edit &&
        (Math.abs(edit.posXPx - start.posXPx) >= 0.5 ||
          Math.abs(edit.posYPx - start.posYPx) >= 0.5 ||
          Math.abs(edit.widthPx - start.widthPx) >= 0.5 ||
          Math.abs(edit.heightPx - start.heightPx) >= 0.5)
      document.body.classList.remove('docs-run-pic-dragging')
      if (!moved) {
        restore()
        return
      }
      // the preview stays until pagination swaps in the re-placed element
      opts.commit(target, {
        posXPx: round1(edit.posXPx),
        posYPx: round1(edit.posYPx),
        widthPx: Math.max(1, Math.round(edit.widthPx)),
        heightPx: Math.max(1, Math.round(edit.heightPx)),
      })
      window.setTimeout(() => {
        if (el.isConnected) restore()
      }, 1500)
    }
    window.addEventListener('mousemove', onMove, true)
    window.addEventListener('mouseup', onUp, true)
  }

  const startMove = (e: MouseEvent, el: HTMLElement, target: HfFloatTarget) =>
    track(
      e,
      el,
      target,
      (dx, dy) => {
        const s = boxOf(el)
        return s ? { ...s, posXPx: s.posXPx + dx, posYPx: s.posYPx + dy } : null
      },
      (edit, start) => {
        el.style.translate = `${edit.posXPx - start.posXPx}px ${edit.posYPx - start.posYPx}px`
      },
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
      (dx, dy) => {
        const s = boxOf(el)
        if (!s) return null
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
      (edit, start) => {
        el.style.width = `${edit.widthPx}px`
        el.style.height = `${edit.heightPx}px`
        el.style.translate = `${edit.posXPx - start.posXPx}px ${edit.posYPx - start.posYPx}px`
      },
      0,
    )
  }

  const onMouseDown = (e: MouseEvent) => {
    if (e.button !== 0 || !opts.editable()) return
    const editing = opts.editing()
    // from the body only a double-click reaches a header picture, and only one
    // drawn in front of the text: a behind-text one leaves the word under it to the body
    if (!editing && e.detail < 2) return
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
      openedAt = performance.now()
      opts.open(target.pageIdx)
      select(target)
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
    root.removeEventListener('mousedown', onMouseDown, true)
    root.removeEventListener('dblclick', onDblClick, true)
    root.removeEventListener('mousemove', onHover)
    window.removeEventListener('keydown', onKey, true)
    root.classList.remove('hf-float-hover')
    frame?.remove()
  }
}
