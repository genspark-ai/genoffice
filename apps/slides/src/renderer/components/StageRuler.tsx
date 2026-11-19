import { useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { PointerEvent as ReactPointerEvent, RefObject } from 'react'
import { computeRulerTicks, type RulerUnit } from '../ruler-ticks'

/**
 * Fixed-chrome stage ruler (PowerPoint-style): lives OUTSIDE the zoomed canvas
 * so its type never scales with zoom, spans the whole work area (ticks continue
 * past the slide edges), and puts 0 at the slide center.
 *
 * - `origin` is where slide coordinate 0 currently falls in ruler-local px
 *   (App measures it from the live .stage-rel rect, so scroll/zoom stay exact).
 * - Press-drag pulls out a guide parallel to the strip (Photoshop/Keynote/
 *   Figma convention): drag DOWN from the horizontal ruler for a horizontal
 *   guide, drag RIGHT from the vertical ruler for a vertical guide.
 * - While the primary button drags anything over the stage, a hairline on each
 *   ruler follows the pointer (imperative transform, no re-render per move).
 */
export function StageRuler({
  vertical = false,
  unit,
  zoom,
  slideLen,
  origin,
  wrapRef,
  onGuidePreview,
  onGuideCommit,
}: {
  vertical?: boolean
  unit: RulerUnit
  zoom: number
  /** slide length along this ruler's axis, unscaled slide px */
  slideLen: number
  /** ruler-local CSS px of slide coordinate 0 */
  origin: number
  /** the scrollable .stage-wrap — pointer source for the position indicator */
  wrapRef: RefObject<HTMLDivElement | null>
  /** pointer positions of a guide pull; App maps them onto the slide's cross axis */
  onGuidePreview: (client: { x: number; y: number }) => void
  onGuideCommit: (client: { x: number; y: number }) => void
}) {
  const rootRef = useRef<HTMLDivElement | null>(null)
  const indicatorRef = useRef<HTMLDivElement | null>(null)
  const [len, setLen] = useState(0)

  // The strip sizes itself with flexbox; tick layout needs the resolved length
  useLayoutEffect(() => {
    const el = rootRef.current
    if (!el) return
    const measure = () => setLen(vertical ? el.clientHeight : el.clientWidth)
    measure()
    const ro = new ResizeObserver(measure)
    ro.observe(el)
    return () => ro.disconnect()
  }, [vertical])

  // Pointer-position hairline while the primary button is held over the stage
  // (object move/resize, marquee, ink, guide drags all qualify)
  useEffect(() => {
    const wrap = wrapRef.current
    if (!wrap) return
    let raf = 0
    let last: { x: number; y: number } | null = null
    const paint = () => {
      raf = 0
      const el = rootRef.current
      const ind = indicatorRef.current
      if (!el || !ind || !last) return
      const r = el.getBoundingClientRect()
      const pos = vertical ? last.y - r.top : last.x - r.left
      ind.style.display = 'block'
      ind.style.transform = vertical ? `translateY(${pos}px)` : `translateX(${pos}px)`
    }
    const hide = () => {
      last = null
      if (raf) cancelAnimationFrame(raf)
      raf = 0
      if (indicatorRef.current) indicatorRef.current.style.display = 'none'
    }
    const move = (e: PointerEvent) => {
      if (!(e.buttons & 1)) {
        if (last) hide()
        return
      }
      last = { x: e.clientX, y: e.clientY }
      if (!raf) raf = requestAnimationFrame(paint)
    }
    wrap.addEventListener('pointermove', move)
    wrap.addEventListener('pointerleave', hide)
    window.addEventListener('pointerup', hide)
    return () => {
      wrap.removeEventListener('pointermove', move)
      wrap.removeEventListener('pointerleave', hide)
      window.removeEventListener('pointerup', hide)
      hide()
    }
  }, [vertical, wrapRef])

  // Press-drag on the strip pulls out a guide PARALLEL to it, dragged away
  // into the canvas (Photoshop/Keynote/Figma convention). The strip only
  // reports pointer positions; App maps them onto the slide's cross axis.
  // A 4px dead zone keeps stray clicks from dropping a guide at the edge.
  const startGuideDrag = (e: ReactPointerEvent<HTMLDivElement>) => {
    if (e.button !== 0) return
    e.preventDefault()
    const el = e.currentTarget
    const start = { x: e.clientX, y: e.clientY }
    let dragging = false
    el.setPointerCapture(e.pointerId)
    const move = (ev: PointerEvent) => {
      if (!dragging) {
        dragging = Math.abs(ev.clientX - start.x) > 4 || Math.abs(ev.clientY - start.y) > 4
        if (!dragging) return
      }
      onGuidePreview({ x: ev.clientX, y: ev.clientY })
    }
    const up = (ev: PointerEvent) => {
      el.removeEventListener('pointermove', move)
      el.removeEventListener('pointerup', up)
      if (dragging) onGuideCommit({ x: ev.clientX, y: ev.clientY })
    }
    el.addEventListener('pointermove', move)
    el.addEventListener('pointerup', up)
  }

  const ticks = computeRulerTicks({ rulerLen: len, origin, slideLen, zoom, unit })
  return (
    <div
      ref={rootRef}
      className={`ruler ${vertical ? 'ruler-v' : 'ruler-h'}`}
      onPointerDown={startGuideDrag}
    >
      {ticks.map((tk) =>
        tk.label != null ? (
          <span
            key={tk.pos}
            className="ruler-num"
            style={vertical ? { top: tk.pos } : { left: tk.pos }}
          >
            {tk.label}
          </span>
        ) : (
          <span
            key={tk.pos}
            className="ruler-tick"
            style={vertical ? { top: tk.pos } : { left: tk.pos }}
          />
        ),
      )}
      <div ref={indicatorRef} className="ruler-indicator" />
    </div>
  )
}
