import React, { useCallback, useRef, useState } from 'react'
import type { PointerEvent as ReactPointerEvent } from 'react'
import type { RenderSlide } from '@genoffice/pptx-render'
import {
  inkNodeHitTest,
  inkNodesOf,
  isPenKind,
  strokeLayers,
  strokePathD,
  type InkPoint,
  type InkStroke,
  type InkTool,
} from './ink'

/**
 * Transparent drawing layer over the slide (pen/highlighter/eraser). Sits above the Konva
 * canvas (inside .stage-rel) and mounts only while a drawing tool is active. Committed
 * strokes are rendered by Konva as regular images; this layer only draws the stroke in
 * progress, and on pen-up passes the whole stroke back to App (rasterize + IPC into the package).
 *
 * Coordinates: overlay size matches RenderSlide (fitWidth viewport px); the outer
 * .stage-scale zoom is divided back out via the getBoundingClientRect ratio.
 */

interface InkOverlayProps {
  slide: RenderSlide
  tool: Exclude<InkTool, 'select'>
  /** hex without '#' */
  color: string
  width: number
  /** Commit a stroke on pen-up (page px coordinates) */
  onCommit: (stroke: InkStroke) => void
  /** Eraser hit: delete these elements (sourceId) */
  onErase: (sourceIds: string[]) => void
}

const ERASER_RADIUS = 6

export function InkOverlay({ slide, tool, color, width, onCommit, onErase }: InkOverlayProps) {
  const overlayRef = useRef<HTMLDivElement>(null)
  const [livePoints, setLivePoints] = useState<InkPoint[] | null>(null)
  const drawingRef = useRef<InkPoint[] | null>(null)
  const erasedRef = useRef<Set<string>>(new Set())

  const localPoint = (e: ReactPointerEvent): InkPoint => {
    const rect = overlayRef.current!.getBoundingClientRect()
    // rect already includes the zoom; map back to RenderSlide px coordinates by ratio
    const scale = rect.width > 0 ? slide.widthPx / rect.width : 1
    return { x: (e.clientX - rect.left) * scale, y: (e.clientY - rect.top) * scale }
  }

  const eraseAt = useCallback(
    (point: InkPoint) => {
      const hits: string[] = []
      for (const entry of inkNodesOf(slide)) {
        if (erasedRef.current.has(entry.node.sourceId)) continue
        if (inkNodeHitTest(entry, point, ERASER_RADIUS)) hits.push(entry.node.sourceId)
      }
      if (hits.length > 0) {
        for (const id of hits) erasedRef.current.add(id)
        onErase(hits)
      }
    },
    [slide, onErase],
  )

  const onPointerDown = (e: ReactPointerEvent) => {
    if (e.button !== 0) return
    // no preventDefault: the ribbon dismisses its popups on bubbled mousedown
    ;(e.target as HTMLElement).setPointerCapture(e.pointerId)
    const point = localPoint(e)
    if (tool === 'eraser') {
      erasedRef.current = new Set()
      drawingRef.current = [] // Non-null marks "eraser drag in progress"
      eraseAt(point)
      return
    }
    drawingRef.current = [point]
    setLivePoints([point])
  }

  const onPointerMove = (e: ReactPointerEvent) => {
    if (!drawingRef.current) return
    const point = localPoint(e)
    if (tool === 'eraser') {
      eraseAt(point)
      return
    }
    const points = drawingRef.current
    const last = points[points.length - 1]!
    if (Math.hypot(point.x - last.x, point.y - last.y) < 0.8) return
    points.push(point)
    setLivePoints([...points])
  }

  const onPointerUp = () => {
    const points = drawingRef.current
    drawingRef.current = null
    if (!isPenKind(tool)) return
    setLivePoints(null)
    if (!points || points.length === 0) return
    onCommit({ tool, color, width, points })
  }

  return (
    <div
      ref={overlayRef}
      className={`ink-overlay ink-active ink-tool-${tool}`}
      onPointerDown={onPointerDown}
      onPointerMove={onPointerMove}
      onPointerUp={onPointerUp}
      onPointerCancel={onPointerUp}
    >
      <svg className="ink-svg" viewBox={`0 0 ${slide.widthPx} ${slide.heightPx}`} aria-hidden>
        {livePoints &&
          isPenKind(tool) &&
          strokeLayers({ tool, color, width, points: livePoints }).map((layer, i) => (
            <path
              key={i}
              d={strokePathD(layer.points)}
              stroke={`#${color}`}
              strokeWidth={layer.width}
              strokeOpacity={layer.alpha}
              fill="none"
              strokeLinecap="round"
              strokeLinejoin="round"
            />
          ))}
      </svg>
    </div>
  )
}
