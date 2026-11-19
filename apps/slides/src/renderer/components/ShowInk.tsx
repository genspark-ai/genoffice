/**
 * Show ink overlay — the presenter view's pen/laser pointer, with the audience window mirroring
 * the same data. Coordinates are normalized 0..1 relative to the slide frame; each side restores
 * them at its own frame size.
 */
import React, { useEffect, useRef } from 'react'

export interface InkStroke {
  color: string
  /** Normalized coordinates alternating x,y */
  points: number[]
}

export function InkLayer({
  strokes,
  laser,
  width,
  height,
}: {
  strokes: InkStroke[]
  laser: { x: number; y: number } | null
  width: number
  height: number
}) {
  const canvasRef = useRef<HTMLCanvasElement>(null)

  useEffect(() => {
    const canvas = canvasRef.current
    if (!canvas) return
    const dpr = window.devicePixelRatio || 1
    canvas.width = Math.round(width * dpr)
    canvas.height = Math.round(height * dpr)
    const ctx = canvas.getContext('2d')
    if (!ctx) return
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0)
    ctx.clearRect(0, 0, width, height)
    ctx.lineCap = 'round'
    ctx.lineJoin = 'round'
    ctx.lineWidth = Math.max(2, width * 0.003)
    for (const s of strokes) {
      if (s.points.length < 4) continue
      ctx.strokeStyle = s.color
      ctx.beginPath()
      ctx.moveTo(s.points[0]! * width, s.points[1]! * height)
      for (let i = 2; i < s.points.length; i += 2) {
        ctx.lineTo(s.points[i]! * width, s.points[i + 1]! * height)
      }
      ctx.stroke()
    }
  }, [strokes, width, height])

  return (
    <div className="ink-layer" style={{ width, height }}>
      <canvas ref={canvasRef} style={{ width, height }} />
      {laser && (
        <div
          className="ink-laser"
          style={{ left: laser.x * width, top: laser.y * height }}
        />
      )}
    </div>
  )
}
