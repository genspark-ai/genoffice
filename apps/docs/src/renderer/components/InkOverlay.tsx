import { useCallback, useEffect, useRef, useState } from 'react'
import type { PointerEvent as ReactPointerEvent } from 'react'
import {
  HIGHLIGHTER_ALPHA,
  nextInkId,
  strokeHitTest,
  strokePathD,
  type InkAnnotation,
  type InkPoint,
  type InkStroke,
  type InkTool,
} from '../editor/ink'

/**
 * Transparent drawing layer covering the page (pen/highlighter/eraser). Lives inside
 * .page-wrap, which is inside the zoomed .doc-zoom container — pointer
 * coordinates are divided by the zoom factor so stroke geometry is stored at
 * 100% zoom, matching the engine's EMU offsets on save.
 *
 * Finished strokes are anchored to the nearest body paragraph (data-idx) and
 * stored relative to its top-left corner, so they follow the paragraph when
 * content above reflows.
 */

interface InkOverlayProps {
  tool: InkTool
  /** hex without '#' */
  color: string
  width: number
  zoom: number
  annotations: InkAnnotation[]
  onAdd: (annotation: InkAnnotation) => void
  onRemove: (ids: string[]) => void
}

/** overlay-relative top-left of every anchored paragraph, at 100% zoom */
type AnchorPositions = Map<number, InkPoint>

const ERASER_RADIUS = 6

function measureAnchors(overlay: HTMLElement, zoomFactor: number): AnchorPositions {
  const positions: AnchorPositions = new Map()
  const base = overlay.getBoundingClientRect()
  const blocks = overlay.parentElement?.querySelectorAll<HTMLElement>('.ProseMirror > [data-idx]')
  for (const el of blocks ? Array.from(blocks) : []) {
    const idx = parseInt(el.dataset.idx ?? '', 10)
    if (!Number.isFinite(idx)) continue
    const rect = el.getBoundingClientRect()
    positions.set(idx, {
      x: (rect.left - base.left) / zoomFactor,
      y: (rect.top - base.top) / zoomFactor,
    })
  }
  return positions
}

/** anchor for a stroke: the paragraph whose top is closest above its first point */
function pickAnchor(positions: AnchorPositions, start: InkPoint): number | null {
  let best: number | null = null
  let bestTop = -Infinity
  for (const [idx, pos] of positions) {
    if (pos.y <= start.y && pos.y > bestTop) {
      bestTop = pos.y
      best = idx
    }
  }
  if (best !== null) return best
  // stroke above every paragraph: fall back to the topmost one
  let top = Infinity
  for (const [idx, pos] of positions) {
    if (pos.y < top) {
      top = pos.y
      best = idx
    }
  }
  return best
}

export function InkOverlay({
  tool,
  color,
  width,
  zoom,
  annotations,
  onAdd,
  onRemove,
}: InkOverlayProps) {
  const overlayRef = useRef<HTMLDivElement>(null)
  const [anchors, setAnchors] = useState<AnchorPositions>(new Map())
  const [livePoints, setLivePoints] = useState<InkPoint[] | null>(null)
  const drawingRef = useRef<InkPoint[] | null>(null)
  const erasedRef = useRef<Set<string>>(new Set())

  const zoomFactor = zoom / 100
  const drawing = tool === 'pen' || tool === 'highlighter'

  const remeasure = useCallback(() => {
    const overlay = overlayRef.current
    if (overlay) setAnchors(measureAnchors(overlay, zoomFactor))
  }, [zoomFactor])

  // paragraph positions shift on every edit/zoom/resize; observe the page and
  // re-measure so anchored strokes stay glued to their paragraphs
  useEffect(() => {
    remeasure()
    const page = overlayRef.current?.parentElement
    if (!page) return
    const observer = new ResizeObserver(remeasure)
    observer.observe(page)
    const mutations = new MutationObserver(remeasure)
    mutations.observe(page, { childList: true, subtree: true, characterData: true })
    return () => {
      observer.disconnect()
      mutations.disconnect()
    }
  }, [remeasure, annotations])

  const localPoint = (e: ReactPointerEvent): InkPoint => {
    const rect = overlayRef.current!.getBoundingClientRect()
    return { x: (e.clientX - rect.left) / zoomFactor, y: (e.clientY - rect.top) / zoomFactor }
  }

  const eraseAt = useCallback(
    (point: InkPoint) => {
      const hits: string[] = []
      for (const ann of annotations) {
        if (erasedRef.current.has(ann.id)) continue
        const pos = anchors.get(ann.anchorIndex)
        if (!pos) continue
        const rel = { x: point.x - pos.x, y: point.y - pos.y }
        if (ann.stroke) {
          if (strokeHitTest(ann.stroke, rel, ERASER_RADIUS)) hits.push(ann.id)
        } else if (ann.image) {
          const { offsetXPx: x, offsetYPx: y, widthPx: w, heightPx: h } = ann.image
          if (rel.x >= x && rel.x <= x + w && rel.y >= y && rel.y <= y + h) hits.push(ann.id)
        }
      }
      if (hits.length > 0) {
        for (const id of hits) erasedRef.current.add(id)
        onRemove(hits)
      }
    },
    [annotations, anchors, onRemove],
  )

  const onPointerDown = (e: ReactPointerEvent) => {
    if (tool === 'select' || e.button !== 0) return
    e.preventDefault()
    ;(e.target as HTMLElement).setPointerCapture(e.pointerId)
    const point = localPoint(e)
    if (tool === 'eraser') {
      erasedRef.current = new Set()
      drawingRef.current = [] // non-null marks "eraser drag active"
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
    const last = points[points.length - 1]
    if (Math.hypot(point.x - last.x, point.y - last.y) < 0.8) return
    points.push(point)
    setLivePoints([...points])
  }

  const onPointerUp = () => {
    const points = drawingRef.current
    drawingRef.current = null
    if (tool !== 'pen' && tool !== 'highlighter') return
    setLivePoints(null)
    if (!points || points.length === 0) return
    const anchorIndex = pickAnchor(anchors, points[0])
    if (anchorIndex === null) return
    const pos = anchors.get(anchorIndex)!
    const stroke: InkStroke = {
      tool,
      color,
      width,
      points: points.map((p) => ({ x: p.x - pos.x, y: p.y - pos.y })),
    }
    onAdd({ id: nextInkId(), anchorIndex, stroke })
  }

  const strokeStyle = (s: InkStroke) => ({
    stroke: `#${s.color}`,
    strokeWidth: s.width,
    strokeOpacity: s.tool === 'highlighter' ? HIGHLIGHTER_ALPHA : 1,
    fill: 'none' as const,
    strokeLinecap: 'round' as const,
    strokeLinejoin: 'round' as const,
  })

  return (
    <div
      ref={overlayRef}
      className={`ink-overlay ${tool !== 'select' ? `ink-active ink-tool-${tool}` : ''}`}
      onPointerDown={onPointerDown}
      onPointerMove={onPointerMove}
      onPointerUp={onPointerUp}
      onPointerCancel={onPointerUp}
    >
      <svg className="ink-svg" aria-hidden>
        {annotations.map((ann) => {
          const pos = anchors.get(ann.anchorIndex)
          if (!pos) return null
          if (ann.stroke) {
            return (
              <path
                key={ann.id}
                d={strokePathD(ann.stroke.points)}
                transform={`translate(${pos.x} ${pos.y})`}
                {...strokeStyle(ann.stroke)}
              />
            )
          }
          if (ann.image?.dataUrl) {
            return (
              <image
                key={ann.id}
                href={ann.image.dataUrl}
                x={pos.x + ann.image.offsetXPx}
                y={pos.y + ann.image.offsetYPx}
                width={ann.image.widthPx}
                height={ann.image.heightPx}
              />
            )
          }
          return null
        })}
        {livePoints && drawing && (
          <path
            d={strokePathD(livePoints)}
            {...strokeStyle({ tool, color, width, points: livePoints })}
          />
        )}
      </svg>
    </div>
  )
}
