import { useRef, useState } from 'react'
import type { PointerEvent as ReactPointerEvent } from 'react'
import { pdfRectToCss, viewToPdf } from './annotations'
import type { PageGeom } from './annotations'
import type { RedactionInput } from '../shared/ipc'

export interface LocalRedaction extends RedactionInput {
  id: string
}

export const REDACTION_MIN_PTS = 3

export function RedactionLayer({
  active,
  geom,
  scale,
  pageWidth,
  pageHeight,
  marks,
  markLabel,
  onCommit,
  onTooSmall,
}: {
  active: boolean
  geom: PageGeom
  scale: number
  pageWidth: number
  pageHeight: number
  marks: LocalRedaction[]
  /** Translated label for pending marks (pass t('redact')); no new locale keys. */
  markLabel: string
  onCommit: (rect: RedactionInput['rect']) => void
  /** Called when a drag is discarded for falling below the minimum size. */
  onTooSmall?: () => void
}) {
  const [live, setLive] = useState<RedactionInput['rect'] | null>(null)
  const start = useRef<[number, number] | null>(null)
  const toPdf = (e: ReactPointerEvent): [number, number] => {
    const box = e.currentTarget.getBoundingClientRect()
    return viewToPdf(geom, (e.clientX - box.left) / scale, (e.clientY - box.top) / scale)
  }
  const down = (e: ReactPointerEvent) => {
    if (!active || e.button !== 0) return
    e.preventDefault()
    e.currentTarget.setPointerCapture(e.pointerId)
    start.current = toPdf(e)
  }
  const move = (e: ReactPointerEvent) => {
    if (!active || !start.current) return
    const [x, y] = toPdf(e)
    const [sx, sy] = start.current
    setLive([Math.min(sx, x), Math.min(sy, y), Math.max(sx, x), Math.max(sy, y)])
  }
  const up = () => {
    const rect = live
    start.current = null
    setLive(null)
    if (!rect) return
    if (rect[2] - rect[0] < REDACTION_MIN_PTS || rect[3] - rect[1] < REDACTION_MIN_PTS) {
      onTooSmall?.()
      return
    }
    onCommit(rect)
  }
  const cancel = () => {
    start.current = null
    setLive(null)
  }
  return (
    <div
      className="pdf-redaction-layer"
      style={{
        width: pageWidth * scale,
        height: pageHeight * scale,
        pointerEvents: active ? 'auto' : 'none',
      }}
      onPointerDown={down}
      onPointerMove={move}
      onPointerUp={up}
      onPointerCancel={cancel}
    >
      {marks.map((mark) => (
        <div
          key={mark.id}
          className="pdf-redaction-mark"
          style={pdfRectToCss(geom, mark.rect, scale)}
          aria-label={markLabel}
        />
      ))}
      {live && (
        <div
          key="live"
          className="pdf-redaction-mark"
          style={pdfRectToCss(geom, live, scale)}
          aria-label={markLabel}
        />
      )}
    </div>
  )
}
