/**
 * Word-parity shape draw mode: picking a shape in the ribbon gallery arms a
 * crosshair on the document instead of inserting immediately. A single click
 * inserts Word's predefined 1x1 inch shape at the click point; a drag draws a
 * custom-size one (holding Shift constrains it to a square, releasing Shift
 * frees it again, Esc cancels). The shape lands where it was drawn through the
 * same floating-object posOffset attrs the move handle uses.
 */
import type { Editor } from '@tiptap/core'
import { shapeClipCss } from '@genoffice/ui'
import { LINE_KINDS } from '@genoffice/docx-engine'
import { isStraightLineKind, shapeBackgroundImage } from './shape-svg'

const EMU_PER_PX = 9525
/** Word's predefined single-click insert size: 1x1 inch. */
export const DEFAULT_SHAPE_EMU = 914400
/** Screen-px distance under which a gesture counts as a click, not a drag. */
const CLICK_THRESHOLD_PX = 3

export interface DrawRectPx {
  x: number
  y: number
  w: number
  h: number
}

/**
 * Resolve a drag (viewport px) into the rectangle to draw. Shift locks a
 * square on the larger delta, growing in the drag direction.
 */
export function resolveDrawRect(
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  shift: boolean,
): DrawRectPx {
  let dx = x2 - x1
  let dy = y2 - y1
  if (shift) {
    const side = Math.max(Math.abs(dx), Math.abs(dy))
    dx = dx < 0 ? -side : side
    dy = dy < 0 ? -side : side
  }
  return {
    x: Math.min(x1, x1 + dx),
    y: Math.min(y1, y1 + dy),
    w: Math.abs(dx),
    h: Math.abs(dy),
  }
}

/**
 * Viewport y where the inserted box's top edge should land. Straight lines
 * collapse to Word's 12px grab band with the stroke at its vertical center,
 * while the ghost previewed the stroke at the drag rect's vertical center —
 * so the band is centered on that line instead of pinned to rect.y.
 */
export function commitTargetY(
  rect: DrawRectPx,
  boxHeightPx: number,
  straightLine: boolean,
): number {
  return straightLine ? rect.y + rect.h / 2 - boxHeightPx / 2 : rect.y
}

/** Drawn rect (viewport px) → shape extent in EMU, floored at 1px so the OOXML stays valid. */
export function drawRectToEmu(
  rect: DrawRectPx,
  zoom: number,
): { widthEmu: number; heightEmu: number } {
  return {
    widthEmu: Math.max(EMU_PER_PX, Math.round((rect.w / zoom) * EMU_PER_PX)),
    heightEmu: Math.max(EMU_PER_PX, Math.round((rect.h / zoom) * EMU_PER_PX)),
  }
}

/**
 * The active-mode cancel lives on window (not module scope): dev HMR swaps the
 * module while a mode is armed, and the stale instance's DOM listeners must
 * still be torn down by the replacement before it arms its own.
 */
const CANCEL_KEY = '__shapeDrawCancel'
// Lazy: keeps the module importable from node-env test suites (no window there)
const cancelStore = (): Record<string, (() => void) | undefined> =>
  window as unknown as Record<string, (() => void) | undefined>
const activeCancel = () => cancelStore()[CANCEL_KEY]

/**
 * Arm the crosshair draw mode on the editor. `insert` performs the actual
 * insertion (widthEmu/heightEmu size, optionally at an explicit doc position)
 * and returns the inserted node's position for the post-insert placement fix.
 * Re-arming (picking another shape) cancels the previous mode.
 */
export function startShapeDrawMode(
  editor: Editor,
  prst: string,
  insert: (opts: { widthEmu: number; heightEmu: number; atPos?: number }) => number | null,
): void {
  activeCancel()?.()
  const view = editor.view
  const dom = view.dom as HTMLElement
  dom.classList.add('doc-shape-drawing')

  let start: { x: number; y: number } | null = null
  let cur: { x: number; y: number } | null = null
  let shift = false
  let ghost: HTMLDivElement | null = null

  // Same zoom handling as the floating-object drag plugin
  const zoomOf = () => {
    const el = document.querySelector('.doc-zoom') as HTMLElement | null
    return el ? parseFloat(getComputedStyle(el).zoom || '1') || 1 : 1
  }

  const isLine = prst in LINE_KINDS

  const updateGhost = () => {
    if (!start || !cur) return
    if (Math.hypot(cur.x - start.x, cur.y - start.y) <= CLICK_THRESHOLD_PX) return
    const r = resolveDrawRect(start.x, start.y, cur.x, cur.y, shift)
    if (!ghost) {
      ghost = document.createElement('div')
      ghost.style.position = 'fixed'
      ghost.style.zIndex = '9999'
      ghost.style.pointerEvents = 'none'
      if (!isLine) {
        // Ghost of the default Office-blue shape the gesture will insert
        ghost.style.background = 'rgba(68,114,196,0.45)'
        ghost.style.border = '1px solid #2F5496'
      }
      ghost.style.boxSizing = 'border-box'
      document.body.appendChild(ghost)
    }
    if (isLine) {
      // Stroke-only preview of the line/connector the gesture will insert
      // (straight kinds land level, so the ghost draws them level too)
      const image = shapeBackgroundImage(
        prst,
        Math.max(8, r.w),
        Math.max(8, r.h),
        undefined,
        '000000',
      )
      ghost.style.backgroundImage = image ?? ''
      ghost.style.backgroundSize = '100% 100%'
      ghost.style.backgroundRepeat = 'no-repeat'
    } else {
      // Recomputed per move: path()-clipped presets are size-dependent
      const clip = shapeClipCss(prst, r.w, r.h)
      ghost.style.clipPath = clip?.clipPath ?? ''
      ghost.style.borderRadius = clip?.borderRadius ?? ''
    }
    ghost.style.left = `${r.x}px`
    ghost.style.top = `${r.y}px`
    ghost.style.width = `${r.w}px`
    ghost.style.height = `${r.h}px`
  }

  const commit = (rect: DrawRectPx, isClick: boolean) => {
    const zoom = zoomOf()
    const { widthEmu, heightEmu } = isClick
      ? { widthEmu: DEFAULT_SHAPE_EMU, heightEmu: DEFAULT_SHAPE_EMU }
      : drawRectToEmu(rect, zoom)
    // Anchor at the top-level block under the gesture origin (caret position otherwise)
    const found = view.posAtCoords({ left: rect.x, top: rect.y })
    let atPos: number | undefined
    if (found) {
      const $p = view.state.doc.resolve(found.pos)
      atPos = $p.depth > 0 ? $p.after(1) : found.pos
    }
    const insertedAt = insert({ widthEmu, heightEmu, atPos })
    if (insertedAt == null) return
    // Two frames so the page layout settles, then nudge the floating shape onto
    // the drawn spot — identical posOffset semantics to dragging the move handle.
    requestAnimationFrame(() =>
      requestAnimationFrame(() => {
        const node = view.state.doc.nodeAt(insertedAt)
        if (!node || node.type.name !== 'docProtected' || !Array.isArray(node.attrs.textboxes))
          return
        const wrapper = view.nodeDOM(insertedAt) as HTMLElement | null
        const box = wrapper?.querySelector('.doc-textbox') as HTMLElement | null
        if (!box) return
        const at = box.getBoundingClientRect()
        const dx = (rect.x - at.left) / zoom
        const dy = (commitTargetY(rect, at.height, isStraightLineKind(prst)) - at.top) / zoom
        if (Math.abs(dx) < 1 && Math.abs(dy) < 1) return
        view.dispatch(
          view.state.tr.setNodeMarkup(insertedAt, undefined, {
            ...node.attrs,
            imageWrap: node.attrs.imageWrap ?? 'square-left',
            imageOffsetXEmu: Number(node.attrs.imageOffsetXEmu ?? 0) + Math.round(dx * EMU_PER_PX),
            imageOffsetYEmu: Number(node.attrs.imageOffsetYEmu ?? 0) + Math.round(dy * EMU_PER_PX),
            imagePosH: null,
            imagePosV: null,
          }),
        )
      }),
    )
  }

  const onMouseDown = (e: MouseEvent) => {
    if (e.button !== 0) return
    e.preventDefault()
    e.stopPropagation()
    start = { x: e.clientX, y: e.clientY }
    cur = { ...start }
    shift = e.shiftKey
    window.addEventListener('mousemove', onWinMove, true)
    window.addEventListener('mouseup', onWinUp, true)
  }
  const onWinMove = (e: MouseEvent) => {
    if (!start) return
    cur = { x: e.clientX, y: e.clientY }
    shift = e.shiftKey
    updateGhost()
  }
  const onWinUp = (e: MouseEvent) => {
    if (!start) return
    const s = start
    const isClick = Math.hypot(e.clientX - s.x, e.clientY - s.y) <= CLICK_THRESHOLD_PX
    const rect = isClick
      ? { x: s.x, y: s.y, w: 0, h: 0 }
      : resolveDrawRect(s.x, s.y, e.clientX, e.clientY, e.shiftKey)
    cleanup()
    commit(rect, isClick)
  }
  const onKey = (e: KeyboardEvent) => {
    if (e.key === 'Escape') {
      cleanup()
      return
    }
    if (e.key === 'Shift' && start) {
      shift = e.type === 'keydown'
      updateGhost()
    }
  }

  const cleanup = () => {
    delete cancelStore()[CANCEL_KEY]
    dom.classList.remove('doc-shape-drawing')
    dom.removeEventListener('mousedown', onMouseDown, true)
    window.removeEventListener('mousemove', onWinMove, true)
    window.removeEventListener('mouseup', onWinUp, true)
    window.removeEventListener('keydown', onKey, true)
    window.removeEventListener('keyup', onKey, true)
    ghost?.remove()
    ghost = null
    start = null
  }

  dom.addEventListener('mousedown', onMouseDown, true)
  window.addEventListener('keydown', onKey, true)
  window.addEventListener('keyup', onKey, true)
  cancelStore()[CANCEL_KEY] = cleanup
}
