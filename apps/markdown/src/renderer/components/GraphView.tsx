import { useCallback, useEffect, useRef, useState } from 'react'
import type { ReactElement } from 'react'
import { useI18n } from '../i18n/locale'
import { layoutGraph, type LayoutPoint } from './graph-layout'
import type { WikiGraph } from '../../shared/ipc'

/** node radius, in world px */
const NODE_R = 7
/** a hit within this many px of a node drags or opens it */
const HIT_R = 14
/** how few nodes still draw their labels */
const LABEL_ALL_BELOW = 120

interface Viewport {
  x: number
  y: number
  scale: number
}

/**
 * The vault's link graph: a force-directed picture of which note links to
 * which, drawn on a canvas (hundreds of nodes redrawn per frame would drown
 * the DOM). The current note wears the accent; dragging a node moves it,
 * dragging the background pans, the wheel zooms, and a click on a node opens
 * it. Nothing here writes to the vault — the graph is a view of the same
 * `[[...]]` links an Obsidian-compatible vault already carries.
 */
export function GraphView({
  current,
  onClose,
  onOpen,
}: {
  current: string | null
  onClose: () => void
  onOpen: (path: string) => void
}): ReactElement {
  const { t } = useI18n()
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const wrapRef = useRef<HTMLDivElement>(null)
  const drawRef = useRef<() => void>(() => {})
  const viewportRef = useRef<Viewport>({ x: 0, y: 0, scale: 1 })
  const layoutRef = useRef<Map<string, LayoutPoint>>(new Map())
  const dragRef = useRef<{
    mode: 'node' | 'pan' | null
    path: string | null
    startX: number
    startY: number
    moved: boolean
  }>({ mode: null, path: null, startX: 0, startY: 0, moved: false })
  const [graph, setGraph] = useState<WikiGraph | null>(null)
  const [failed, setFailed] = useState(false)

  useEffect(() => {
    let cancelled = false
    // The browser test harness has no Electron channel; an absent API is an
    // empty graph, not a failed one.
    if (typeof window.markdownApi?.wikiGraph !== 'function') {
      setGraph({ nodes: [], edges: [], truncated: false, totalNotes: 0 })
      return
    }
    window.markdownApi
      .wikiGraph()
      .then((g) => {
        if (!cancelled) setGraph(g)
      })
      .catch(() => {
        if (!cancelled) setFailed(true)
      })
    return () => {
      cancelled = true
    }
  }, [])

  /** theme colors, read off the canvas's own computed style at draw time */
  const colors = useCallback((): { text: string; dim: string; accent: string; line: string } => {
    const el = canvasRef.current
    const style = el ? getComputedStyle(el) : null
    const pick = (name: string, fallback: string): string =>
      style?.getPropertyValue(name).trim() || fallback
    return {
      text: pick('--text', '#333333'),
      dim: pick('--text-tertiary', '#999999'),
      accent: pick('--accent', '#4a7d00'),
      line: pick('--border-strong', '#cccccc'),
    }
  }, [])

  useEffect(() => {
    const canvas = canvasRef.current
    const wrap = wrapRef.current
    if (!canvas || !wrap || !graph) return
    const w = wrap.clientWidth
    const h = wrap.clientHeight
    if (w === 0 || h === 0) return
    const dpr = window.devicePixelRatio || 1
    canvas.width = Math.round(w * dpr)
    canvas.height = Math.round(h * dpr)

    layoutRef.current = layoutGraph(graph.nodes, graph.edges, w, h)
    const c = colors()

    const draw = (): void => {
      const ctx = canvas.getContext('2d')
      if (!ctx) return
      const vp = viewportRef.current
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0)
      ctx.clearRect(0, 0, w, h)
      ctx.save()
      ctx.translate(vp.x, vp.y)
      ctx.scale(vp.scale, vp.scale)
      const labelEvery = graph.nodes.length <= LABEL_ALL_BELOW || vp.scale > 1.4
      ctx.strokeStyle = c.line
      ctx.lineWidth = 1 / vp.scale
      ctx.globalAlpha = 0.55
      ctx.beginPath()
      for (const edge of graph.edges) {
        const a = layoutRef.current.get(edge.from)
        const b = layoutRef.current.get(edge.to)
        if (!a || !b) continue
        ctx.moveTo(a.x, a.y)
        ctx.lineTo(b.x, b.y)
      }
      ctx.stroke()
      ctx.globalAlpha = 1
      for (const node of graph.nodes) {
        const p = layoutRef.current.get(node.path)
        if (!p) continue
        const isCurrent = node.path === current
        ctx.beginPath()
        ctx.arc(p.x, p.y, isCurrent ? NODE_R + 2 : NODE_R, 0, Math.PI * 2)
        ctx.fillStyle = isCurrent ? c.accent : c.dim
        ctx.fill()
        if (isCurrent || labelEvery) {
          ctx.fillStyle = isCurrent ? c.text : c.dim
          ctx.font = `${isCurrent ? '600 ' : ''}11px -apple-system, system-ui, sans-serif`
          ctx.textAlign = 'center'
          ctx.fillText(node.name, p.x, p.y + NODE_R + 13)
        }
      }
      ctx.restore()
    }
    drawRef.current = draw
    draw()
  }, [graph, current, colors])

  const nodeAt = useCallback((clientX: number, clientY: number): string | null => {
    const canvas = canvasRef.current
    if (!canvas) return null
    const rect = canvas.getBoundingClientRect()
    const vp = viewportRef.current
    const x = (clientX - rect.left - vp.x) / vp.scale
    const y = (clientY - rect.top - vp.y) / vp.scale
    let best: string | null = null
    let bestD = HIT_R
    for (const [path, p] of layoutRef.current) {
      const d = Math.hypot(p.x - x, p.y - y)
      if (d <= bestD) {
        bestD = d
        best = path
      }
    }
    return best
  }, [])

  const onPointerDown = (e: React.PointerEvent<HTMLCanvasElement>): void => {
    e.currentTarget.setPointerCapture(e.pointerId)
    const path = nodeAt(e.clientX, e.clientY)
    dragRef.current = {
      mode: path ? 'node' : 'pan',
      path,
      startX: e.clientX,
      startY: e.clientY,
      moved: false,
    }
  }

  const onPointerMove = (e: React.PointerEvent<HTMLCanvasElement>): void => {
    const drag = dragRef.current
    if (!drag.mode) return
    const dx = e.clientX - drag.startX
    const dy = e.clientY - drag.startY
    if (Math.abs(dx) + Math.abs(dy) > 3) drag.moved = true
    if (drag.mode === 'pan') {
      viewportRef.current = {
        ...viewportRef.current,
        x: viewportRef.current.x + dx,
        y: viewportRef.current.y + dy,
      }
      drag.startX = e.clientX
      drag.startY = e.clientY
      drawRef.current()
      return
    }
    if (drag.mode === 'node' && drag.path && drag.moved) {
      const canvas = canvasRef.current
      if (!canvas) return
      const rect = canvas.getBoundingClientRect()
      const vp = viewportRef.current
      const p = layoutRef.current.get(drag.path)
      if (p) {
        p.x = (e.clientX - rect.left - vp.x) / vp.scale
        p.y = (e.clientY - rect.top - vp.y) / vp.scale
      }
      drawRef.current()
    }
  }

  const onPointerUp = (e: React.PointerEvent<HTMLCanvasElement>): void => {
    const drag = dragRef.current
    dragRef.current = { mode: null, path: null, startX: 0, startY: 0, moved: false }
    // a press that never became a drag is a click: open the node
    if (drag.mode === 'node' && !drag.moved && drag.path) onOpen(drag.path)
    e.currentTarget.releasePointerCapture(e.pointerId)
  }

  const onWheel = (e: React.WheelEvent<HTMLCanvasElement>): void => {
    const factor = e.deltaY < 0 ? 1.1 : 1 / 1.1
    const v = viewportRef.current
    const scale = Math.min(6, Math.max(0.2, v.scale * factor))
    // keep the point under the cursor stationary while zooming
    const rect = e.currentTarget.getBoundingClientRect()
    const mx = e.clientX - rect.left
    const my = e.clientY - rect.top
    const k = scale / v.scale
    viewportRef.current = { scale, x: mx - (mx - v.x) * k, y: my - (my - v.y) * k }
    drawRef.current()
  }

  // The overlay covers the ribbon strip, whose `-webkit-app-region: drag`
  // swallows the presses the header (the close button among them) needs —
  // `no-drag` on the button alone does not win while the strip keeps its
  // region (the popover-open row flip exists for the same reason). Mark the
  // open graph on <html> so the row yields, and dismiss on Escape the way a
  // full-screen view elsewhere does (RedactMenu).
  useEffect(() => {
    document.documentElement.classList.add('markdown-graph-open')
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', onKey)
    return () => {
      document.documentElement.classList.remove('markdown-graph-open')
      window.removeEventListener('keydown', onKey)
    }
  }, [onClose])

  return (
    <div className="graph-overlay" role="dialog" aria-label={t('graph')}>
      <div className="graph-head">
        <span className="graph-title">{t('graph')}</span>
        {graph?.truncated === true && (
          <span className="graph-truncated">
            {t('graphTruncated', { shown: graph.nodes.length, total: graph.totalNotes })}
          </span>
        )}
        <button
          type="button"
          className="graph-close"
          aria-label={t('graphClose')}
          onClick={onClose}
        >
          ✕
        </button>
      </div>
      <div className="graph-body" ref={wrapRef}>
        {failed || (graph !== null && graph.nodes.length === 0) ? (
          <div className="graph-empty">{failed ? t('slashNoResults') : t('graphEmpty')}</div>
        ) : graph !== null ? (
          <canvas
            ref={canvasRef}
            className="graph-canvas"
            onPointerDown={onPointerDown}
            onPointerMove={onPointerMove}
            onPointerUp={onPointerUp}
            onWheel={onWheel}
          />
        ) : null}
      </div>
    </div>
  )
}
