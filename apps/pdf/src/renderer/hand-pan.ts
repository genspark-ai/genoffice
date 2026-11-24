import { useEffect, type RefObject } from 'react'

/** Pointer travel before a press turns into a pan, so plain clicks (links, pins) still land */
const DRAG_THRESHOLD = 4

/** Presses on these keep their own behavior under the hand tool */
const PASS_THROUGH =
  'input, textarea, select, button, [contenteditable="true"], .pdf-draw-shape, .pdf-note-pin, .pdf-note-margin, .pdf-del-popup, .pdf-imgedit-layer'

export const PANNING_CLASS = 'pdf-panning'

/** Drag-to-scroll on `el`; returns the teardown */
export function attachHandPan(el: HTMLElement): () => void {
  let start: { x: number; y: number; left: number; top: number; id: number } | null = null
  let panning = false

  const onDown = (e: PointerEvent) => {
    if (e.button !== 0) return
    const target = e.target as Element
    if (target.closest?.(PASS_THROUGH)) return
    // a press on the container's own scrollbar
    const box = el.getBoundingClientRect()
    if (e.clientX - box.left >= el.clientWidth || e.clientY - box.top >= el.clientHeight) return
    start = { x: e.clientX, y: e.clientY, left: el.scrollLeft, top: el.scrollTop, id: e.pointerId }
    panning = false
  }
  const onMove = (e: PointerEvent) => {
    if (!start || e.pointerId !== start.id) return
    const dx = e.clientX - start.x
    const dy = e.clientY - start.y
    if (!panning) {
      if (Math.hypot(dx, dy) < DRAG_THRESHOLD) return
      panning = true
      el.classList.add(PANNING_CLASS)
      try {
        el.setPointerCapture(start.id)
      } catch {
        // pointer already gone: window listeners still end the gesture
      }
    }
    el.scrollLeft = start.left - dx
    el.scrollTop = start.top - dy
  }
  const end = (e: PointerEvent) => {
    if (!start || e.pointerId !== start.id) return
    start = null
    el.classList.remove(PANNING_CLASS)
    if (!panning) return
    // the click that follows a drag must not activate what was under the pointer
    const swallow = (ev: MouseEvent) => {
      ev.stopPropagation()
      ev.preventDefault()
    }
    window.addEventListener('click', swallow, { capture: true, once: true })
    setTimeout(() => window.removeEventListener('click', swallow, { capture: true }), 0)
  }

  el.addEventListener('pointerdown', onDown)
  window.addEventListener('pointermove', onMove)
  window.addEventListener('pointerup', end)
  window.addEventListener('pointercancel', end)
  return () => {
    el.removeEventListener('pointerdown', onDown)
    window.removeEventListener('pointermove', onMove)
    window.removeEventListener('pointerup', end)
    window.removeEventListener('pointercancel', end)
    el.classList.remove(PANNING_CLASS)
  }
}

export function useHandPan(ref: RefObject<HTMLElement | null>, active: boolean): void {
  useEffect(() => {
    const el = ref.current
    if (!active || !el) return
    return attachHandPan(el)
  }, [ref, active])
}
