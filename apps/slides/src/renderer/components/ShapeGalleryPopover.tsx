/**
 * Floating shape gallery for the "Change Shape" context-menu entry (WPS/PowerPoint
 * parity): opens at the click point, kept within the viewport; closes on outside
 * click / Escape / scroll. Picking a preset swaps the shape's geometry only.
 */
import React, { useEffect, useLayoutEffect, useRef, useState } from 'react'
import { isLineDrawKind } from '../draw-shape'
import { SHAPE_GALLERY } from '../insert-presets'
import { ShapePreview } from './gallery-previews'

// Line entries are insert-only: their kinds are aliases resolved at insert
// time (not valid <a:prstGeom> presets), and swapping a filled shape into a
// connector is not a geometry change. PowerPoint's Change Shape gallery
// excludes lines for the same reason. The original group objects are kept
// (not spread): their `group` title is a live getter that follows the UI
// language, so it must be read at render time.
const CHANGE_SHAPE_GALLERY = SHAPE_GALLERY.map((group) => ({
  group,
  shapes: group.shapes.filter((s) => !isLineDrawKind(s.prst)),
})).filter((entry) => entry.shapes.length > 0)

interface Props {
  x: number
  y: number
  onPick: (prst: string) => void
  onClose: () => void
}

/** Shared gallery body used by both the canvas context menu and Shape Tools ribbon. */
export function ShapeGalleryContent({ onPick }: Pick<Props, 'onPick'>) {
  return (
    <div className="rb-shape-gallery">
      {CHANGE_SHAPE_GALLERY.map(({ group, shapes }, index) => (
        <div key={index}>
          <div className="rb-drop-title">{group.group}</div>
          <div className="rb-shape-grid">
            {shapes.map((s) => (
              <button
                key={s.prst}
                className="rb-shape-cell"
                data-prst={s.prst}
                data-tip={s.label}
                aria-label={s.label}
                onClick={() => onPick(s.prst)}
              >
                <ShapePreview prst={s.prst} size={18} />
              </button>
            ))}
          </div>
        </div>
      ))}
    </div>
  )
}

export function ShapeGalleryPopover({ x, y, onPick, onClose }: Props) {
  const ref = useRef<HTMLDivElement>(null)
  const [pos, setPos] = useState({ x, y })

  useLayoutEffect(() => {
    const el = ref.current
    if (!el) return
    const { innerWidth, innerHeight } = window
    const r = el.getBoundingClientRect()
    setPos({
      x: Math.min(x, Math.max(0, innerWidth - r.width - 4)),
      y: Math.min(y, Math.max(0, innerHeight - r.height - 4)),
    })
  }, [x, y])

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    const onDown = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) onClose()
    }
    window.addEventListener('keydown', onKey)
    window.addEventListener('mousedown', onDown)
    window.addEventListener('blur', onClose)
    return () => {
      window.removeEventListener('keydown', onKey)
      window.removeEventListener('mousedown', onDown)
      window.removeEventListener('blur', onClose)
    }
  }, [onClose])

  return (
    <div
      ref={ref}
      className="ctx-menu shape-gallery-pop"
      style={{ left: pos.x, top: pos.y }}
      onContextMenu={(e) => e.preventDefault()}
    >
      <ShapeGalleryContent
        onPick={(prst) => {
          onClose()
          onPick(prst)
        }}
      />
    </div>
  )
}
