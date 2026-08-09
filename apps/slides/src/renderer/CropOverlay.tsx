/**
 * Picture crop mode overlay.
 *
 * Entry: context menu "Crop picture" → receives the picture node's box + current srcRect.
 * Interaction: drag 8 handles to adjust the crop frame; Enter / click outside to confirm; Esc to cancel.
 * Output: onConfirm({ l, t, r, b }) or onCancel().
 *
 * Coordinate system: everything in the Konva Stage's CSS px (fitWidth viewport).
 * Non-destructive: the full original image is shown ghosted behind the crop frame
 * (reconstructed from the current srcRect), and the frame may be dragged OUTWARD
 * up to the original bounds to restore cropped-away content. The confirmed
 * l/t/r/b are absolute 0..1 fractions of the full original image; null = frame
 * covers the whole original (crop removed).
 */
import React, { useCallback, useEffect, useRef, useState } from 'react'

export interface CropRect {
  l: number
  t: number
  r: number
  b: number
}

export interface NodeBox {
  x: number
  y: number
  w: number
  h: number
}

interface Props {
  /** The picture element's current (cropped) pixel box — the frame's start position */
  box: NodeBox
  /** Full original-image extent; the frame is clamped to this, not to `box` */
  fullBox: NodeBox
  /** Source bitmap for the ghost preview (crop area beyond the current view) */
  imgSrc?: string
  /** Absolute fractions of the full original image; null = crop removed */
  onConfirm: (rect: CropRect | null) => void
  onCancel: () => void
}

type HandleId = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w'

const HANDLE_SIZE = 8
const HANDLE_HALF = HANDLE_SIZE / 2

/**
 * Swallow the next `click` the browser synthesizes from the current mousedown.
 * Lives outside the component: `confirm()` unmounts the overlay before the
 * mouseup/click fire, so an effect-scoped listener would be removed by the
 * cleanup and never see the click it is meant to block.
 */
function swallowNextClick() {
  const timer = window.setTimeout(() => {
    window.removeEventListener('click', swallow, { capture: true })
  }, 400)
  function swallow(ev: MouseEvent) {
    window.clearTimeout(timer)
    ev.preventDefault()
    ev.stopPropagation()
  }
  window.addEventListener('click', swallow, { capture: true, once: true })
}

/** PowerPoint-style crop cursor for the edge-midpoint handles: a T-bar (⊥) whose
 * stem points outward from the picture; dragging pushes that edge inward. */
function cropCursor(rotation: number, fallback: string): string {
  const svg =
    '<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24">' +
    `<g transform="rotate(${rotation} 12 12)">` +
    '<path d="M9 3h6v9h5v5H4v-5h5z" fill="#000" stroke="#fff" stroke-width="1.5" stroke-linejoin="round"/>' +
    '</g></svg>'
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}") 10 10, ${fallback}`
}

const CURSOR: Record<HandleId, string> = {
  nw: 'nwse-resize',
  n: cropCursor(0, 'ns-resize'),
  ne: 'nesw-resize',
  e: cropCursor(90, 'ew-resize'),
  se: 'nwse-resize',
  s: cropCursor(180, 'ns-resize'),
  sw: 'nesw-resize',
  w: cropCursor(270, 'ew-resize'),
}

/**
 * Positions of the 8 handles relative to the crop frame's top-left
 */
function handlePositions(cx: number, cy: number, cw: number, ch: number) {
  const mx = cx + cw / 2
  const my = cy + ch / 2
  const ex = cx + cw
  const ey = cy + ch
  return {
    nw: { x: cx, y: cy },
    n: { x: mx, y: cy },
    ne: { x: ex, y: cy },
    e: { x: ex, y: my },
    se: { x: ex, y: ey },
    s: { x: mx, y: ey },
    sw: { x: cx, y: ey },
    w: { x: cx, y: my },
  } as Record<HandleId, { x: number; y: number }>
}

/**
 * Pixel crop frame → absolute crop ratios 0..1 of the full original image
 */
function srcRectFromRect(full: NodeBox, cx: number, cy: number, cw: number, ch: number): CropRect {
  const l = (cx - full.x) / full.w
  const t = (cy - full.y) / full.h
  const r = (full.x + full.w - cx - cw) / full.w
  const b = (full.y + full.h - cy - ch) / full.h
  // the frame is clamped inside `full` with a 20px minimum, so each fraction is
  // naturally in [0, 1) and opposite edges sum below 1 - only guard rounding noise
  const clamp = (v: number) => Math.max(0, v)
  return { l: clamp(l), t: clamp(t), r: clamp(r), b: clamp(b) }
}

export function CropOverlay({ box, fullBox, imgSrc, onConfirm, onCancel }: Props) {
  // The frame starts at the current visible region; the ghost of the full original
  // renders behind it, so dragging outward reveals (and restores) cropped content.
  const [cx, setCx] = useState(box.x)
  const [cy, setCy] = useState(box.y)
  const [cw, setCw] = useState(box.w)
  const [ch, setCh] = useState(box.h)

  const dragRef = useRef<{
    handle: HandleId
    startMouseX: number
    startMouseY: number
    startCx: number
    startCy: number
    startCw: number
    startCh: number
  } | null>(null)

  const onMouseDown = useCallback(
    (handle: HandleId, e: React.MouseEvent) => {
      e.stopPropagation()
      e.preventDefault()
      dragRef.current = {
        handle,
        startMouseX: e.clientX,
        startMouseY: e.clientY,
        startCx: cx,
        startCy: cy,
        startCw: cw,
        startCh: ch,
      }
    },
    [cx, cy, cw, ch],
  )

  useEffect(() => {
    const onMove = (e: MouseEvent) => {
      const d = dragRef.current
      if (!d) return
      const dx = e.clientX - d.startMouseX
      const dy = e.clientY - d.startMouseY
      const MIN = 20
      let nx = d.startCx,
        ny = d.startCy,
        nw = d.startCw,
        nh = d.startCh
      switch (d.handle) {
        case 'nw':
          nx = Math.min(d.startCx + dx, d.startCx + d.startCw - MIN)
          ny = Math.min(d.startCy + dy, d.startCy + d.startCh - MIN)
          nw = d.startCw - (nx - d.startCx)
          nh = d.startCh - (ny - d.startCy)
          break
        case 'n':
          ny = Math.min(d.startCy + dy, d.startCy + d.startCh - MIN)
          nh = d.startCh - (ny - d.startCy)
          break
        case 'ne':
          ny = Math.min(d.startCy + dy, d.startCy + d.startCh - MIN)
          nh = d.startCh - (ny - d.startCy)
          nw = Math.max(d.startCw + dx, MIN)
          break
        case 'e':
          nw = Math.max(d.startCw + dx, MIN)
          break
        case 'se':
          nw = Math.max(d.startCw + dx, MIN)
          nh = Math.max(d.startCh + dy, MIN)
          break
        case 's':
          nh = Math.max(d.startCh + dy, MIN)
          break
        case 'sw':
          nx = Math.min(d.startCx + dx, d.startCx + d.startCw - MIN)
          nw = d.startCw - (nx - d.startCx)
          nh = Math.max(d.startCh + dy, MIN)
          break
        case 'w':
          nx = Math.min(d.startCx + dx, d.startCx + d.startCw - MIN)
          nw = d.startCw - (nx - d.startCx)
          break
      }
      // Clamp within the ORIGINAL image bounds — outward drag restores cropped content
      nx = Math.max(fullBox.x, nx)
      ny = Math.max(fullBox.y, ny)
      nw = Math.min(nw, fullBox.x + fullBox.w - nx)
      nh = Math.min(nh, fullBox.y + fullBox.h - ny)
      setCx(nx)
      setCy(ny)
      setCw(nw)
      setCh(nh)
    }
    const onUp = () => {
      dragRef.current = null
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
    return () => {
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
    }
  }, [fullBox])

  // Keyboard: Enter confirm, Esc cancel
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Enter') {
        e.preventDefault()
        const r = srcRectFromRect(fullBox, cx, cy, cw, ch)
        const isEmpty = !r.l && !r.t && !r.r && !r.b
        onConfirm(isEmpty ? null : r)
      } else if (e.key === 'Escape') {
        e.preventDefault()
        onCancel()
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [fullBox, cx, cy, cw, ch, onConfirm, onCancel])

  const confirm = useCallback(() => {
    const r = srcRectFromRect(fullBox, cx, cy, cw, ch)
    const isEmpty = !r.l && !r.t && !r.r && !r.b
    onConfirm(isEmpty ? null : r)
  }, [fullBox, cx, cy, cw, ch, onConfirm])

  // The overlay only covers the slide canvas; a mousedown anywhere beyond it
  // (thumbnails, ribbon, the gray backdrop) also confirms — capture phase so the
  // crop lands before whatever was clicked reacts to its own event.
  const rootRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    const onDown = (e: MouseEvent) => {
      const root = rootRef.current
      if (root && e.target instanceof Node && root.contains(e.target)) return
      if (e.button !== 0) return // primary button only (right-click = context menu)
      // The first outside press only confirms: swallow it (and the click it
      // synthesizes) so the control under the cursor does not also act - e.g. Undo
      // racing the in-flight crop IPC, or the Crop button re-entering crop mode
      // from pre-commit geometry. The click swallower is registered outside this
      // effect so unmounting the overlay cannot cancel it.
      e.preventDefault()
      e.stopPropagation()
      swallowNextClick()
      confirm()
    }
    window.addEventListener('mousedown', onDown, true)
    return () => {
      window.removeEventListener('mousedown', onDown, true)
    }
  }, [confirm])

  const handles = handlePositions(cx, cy, cw, ch)

  return (
    <div
      ref={rootRef}
      className="crop-overlay"
      onMouseDown={(e) => {
        // Clicking outside the handles = confirm
        e.stopPropagation()
        confirm()
      }}
      style={{ position: 'absolute', inset: 0, cursor: 'default' }}
    >
      {/* Ghost of the FULL original image: crisp inside the frame (masks dim the rest) */}
      {imgSrc && (
        <img
          src={imgSrc}
          alt=""
          draggable={false}
          style={{
            position: 'absolute',
            left: fullBox.x,
            top: fullBox.y,
            width: fullBox.w,
            height: fullBox.h,
            pointerEvents: 'none',
          }}
        />
      )}

      {/* Mask layer: translucent outside the crop frame, over the whole original */}
      {/* top */}
      <div
        className="crop-mask"
        style={{ left: fullBox.x, top: fullBox.y, width: fullBox.w, height: cy - fullBox.y }}
      />
      {/* bottom */}
      <div
        className="crop-mask"
        style={{
          left: fullBox.x,
          top: cy + ch,
          width: fullBox.w,
          height: fullBox.y + fullBox.h - cy - ch,
        }}
      />
      {/* left (middle strip) */}
      <div
        className="crop-mask"
        style={{ left: fullBox.x, top: cy, width: cx - fullBox.x, height: ch }}
      />
      {/* right (middle strip) */}
      <div
        className="crop-mask"
        style={{ left: cx + cw, top: cy, width: fullBox.x + fullBox.w - cx - cw, height: ch }}
      />

      {/* Crop frame border */}
      <div
        className="crop-frame"
        style={{ left: cx, top: cy, width: cw, height: ch }}
        onMouseDown={(e) => e.stopPropagation()}
      />

      {/* The 8 handles */}
      {(Object.entries(handles) as [HandleId, { x: number; y: number }][]).map(([id, pos]) => (
        <div
          key={id}
          className="crop-handle"
          style={{
            left: pos.x - HANDLE_HALF,
            top: pos.y - HANDLE_HALF,
            width: HANDLE_SIZE,
            height: HANDLE_SIZE,
            cursor: CURSOR[id],
          }}
          onMouseDown={(e) => onMouseDown(id, e)}
        />
      ))}
    </div>
  )
}
