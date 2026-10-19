import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { MindMapTreeNode } from '@genoffice/agent-core'
import {
  countMindMapNodes,
  layoutMindMap,
  MIND_MAP_EXPORT_COLORS,
  MIND_MAP_METRICS,
} from './mind-map'
import { useI18n } from '../i18n/locale'

/**
 * Interactive mind map card rendered inside an assistant message (the
 * NotebookLM-style diagram): a self-drawn SVG tidy tree with collapsible
 * branches, wheel zoom + drag pan, node-click-to-ask, block-citation jumps
 * and PNG export. No diagram library — plain SVG keeps the bundle flat and
 * gives full control over node-level interaction.
 */

const ZOOM_MIN = 0.3
const ZOOM_MAX = 2.5

/** Font stack baked into exported SVGs (the live DOM uses the stylesheet). */
const EXPORT_FONT = "'Arial', 'Inter', -apple-system, 'PingFang SC', 'Microsoft YaHei', sans-serif"

interface MindMapCardProps {
  tree: MindMapTreeNode
  /** click a node: ask the assistant about that topic */
  onAsk: (label: string) => void
  /** click a node's block badge: jump to the cited document block */
  onJumpBlock: (index: number) => void
}

export function MindMapCard({ tree, onAsk, onJumpBlock }: MindMapCardProps) {
  const { t } = useI18n()
  const [collapsed, setCollapsed] = useState<ReadonlySet<string>>(() => new Set())
  const [zoom, setZoom] = useState(1)
  const [pan, setPan] = useState({ x: 8, y: 8 })
  const viewportRef = useRef<HTMLDivElement>(null)
  const svgRef = useRef<SVGSVGElement>(null)
  const zoomRef = useRef(1)
  const dragRef = useRef<{ id: number; x: number; y: number } | null>(null)

  const layout = useMemo(() => layoutMindMap(tree, collapsed), [tree, collapsed])
  const totalNodes = useMemo(() => countMindMapNodes(tree), [tree])

  const fitView = useCallback(() => {
    const vp = viewportRef.current
    zoomRef.current = 1
    setZoom(1)
    setPan({ x: 8, y: vp ? Math.max(8, (vp.clientHeight - layout.height) / 2) : 8 })
  }, [layout.height])

  // fit once per new map (collapse toggles keep the user's view)
  const lastTreeRef = useRef<MindMapTreeNode | null>(null)
  useEffect(() => {
    if (lastTreeRef.current === tree) return
    lastTreeRef.current = tree
    fitView()
  }, [tree, fitView])

  // Ctrl/Cmd+wheel zooms around the cursor (Chromium reports a trackpad pinch
  // as ctrl+wheel) — the same gesture the editor canvas uses. A plain wheel is
  // left alone so the chat panel still scrolls past the card; panning is drag.
  useEffect(() => {
    const vp = viewportRef.current
    if (!vp) return
    const onWheel = (e: WheelEvent) => {
      if (!e.ctrlKey && !e.metaKey) return
      e.preventDefault()
      const rect = vp.getBoundingClientRect()
      const ox = e.clientX - rect.left
      const oy = e.clientY - rect.top
      const next = Math.min(
        ZOOM_MAX,
        Math.max(ZOOM_MIN, zoomRef.current * Math.exp(-e.deltaY * 0.0015)),
      )
      const k = next / zoomRef.current
      zoomRef.current = next
      setZoom(next)
      setPan((p) => ({ x: ox - (ox - p.x) * k, y: oy - (oy - p.y) * k }))
    }
    vp.addEventListener('wheel', onWheel, { passive: false })
    return () => vp.removeEventListener('wheel', onWheel)
  }, [])

  const toggleCollapse = (id: string) =>
    setCollapsed((prev) => {
      const next = new Set(prev)
      if (next.has(id)) next.delete(id)
      else next.add(id)
      return next
    })

  const onPointerDown = (e: React.PointerEvent<HTMLDivElement>) => {
    // dragging starts on the background only; nodes handle their own clicks
    if ((e.target as Element).closest('.ai-mindmap-node')) return
    dragRef.current = { id: e.pointerId, x: e.clientX, y: e.clientY }
    e.currentTarget.setPointerCapture(e.pointerId)
  }
  const onPointerMove = (e: React.PointerEvent<HTMLDivElement>) => {
    const d = dragRef.current
    if (!d || d.id !== e.pointerId) return
    setPan((p) => ({ x: p.x + (e.clientX - d.x), y: p.y + (e.clientY - d.y) }))
    dragRef.current = { ...d, x: e.clientX, y: e.clientY }
  }
  const onPointerUp = (e: React.PointerEvent<HTMLDivElement>) => {
    if (dragRef.current?.id === e.pointerId) dragRef.current = null
  }

  /** Serialize the SVG, bake in the fixed export palette, rasterize 2x, download. */
  const exportPng = () => {
    const svg = svgRef.current
    if (!svg) return
    const { width, height } = layout
    const clone = svg.cloneNode(true) as SVGSVGElement
    clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg')
    clone.setAttribute('font-family', EXPORT_FONT)
    const bg = document.createElementNS('http://www.w3.org/2000/svg', 'rect')
    bg.setAttribute('x', '0')
    bg.setAttribute('y', '0')
    bg.setAttribute('width', String(width))
    bg.setAttribute('height', String(height))
    bg.setAttribute('fill', MIND_MAP_EXPORT_COLORS.background)
    clone.insertBefore(bg, clone.firstChild)
    const c = MIND_MAP_EXPORT_COLORS
    for (const el of Array.from(clone.querySelectorAll('[data-export]'))) {
      const kind = el.getAttribute('data-export')
      const tag = el.tagName.toLowerCase()
      const set = (attrs: Record<string, string>) => {
        for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, v)
      }
      if (kind === 'edge' && tag === 'path') {
        set({ fill: 'none', stroke: c.edge, 'stroke-width': '1.5' })
      } else if (kind === 'node-box') {
        set({ fill: c.nodeFill, stroke: c.nodeStroke, 'stroke-width': '1' })
      } else if (kind === 'root-box') {
        set({ fill: c.rootFill, stroke: c.rootStroke, 'stroke-width': '1' })
      } else if (kind === 'node-text') {
        set({ fill: c.nodeText, 'font-size': '12.5' })
      } else if (kind === 'root-text') {
        set({ fill: c.rootText, 'font-size': '12.5' })
      } else if (kind === 'badge') {
        set({ fill: c.badgeFill })
      } else if (kind === 'badge-text') {
        set({ fill: c.badgeText, 'font-size': '10' })
      } else if (kind === 'toggle') {
        set({ fill: c.toggleFill, stroke: c.toggleStroke, 'stroke-width': '1' })
      } else if (kind === 'toggle-text') {
        set({ fill: c.toggleText, 'font-size': '10' })
      }
    }
    const raw = new XMLSerializer().serializeToString(clone)
    const svgUrl = URL.createObjectURL(new Blob([raw], { type: 'image/svg+xml;charset=utf-8' }))
    const img = new Image()
    img.onload = () => {
      URL.revokeObjectURL(svgUrl)
      const scale = 2
      const canvas = document.createElement('canvas')
      canvas.width = Math.round(width * scale)
      canvas.height = Math.round(height * scale)
      const ctx = canvas.getContext('2d')
      if (!ctx) return
      ctx.scale(scale, scale)
      ctx.drawImage(img, 0, 0)
      canvas.toBlob((blob) => {
        if (!blob) return
        const pngUrl = URL.createObjectURL(blob)
        const a = document.createElement('a')
        a.href = pngUrl
        a.download = 'mindmap.png'
        a.click()
        URL.revokeObjectURL(pngUrl)
      }, 'image/png')
    }
    img.onerror = () => URL.revokeObjectURL(svgUrl)
    img.src = svgUrl
  }

  const { nodeWidth: W, nodeHeight: H } = MIND_MAP_METRICS

  return (
    <div className="ai-mindmap">
      <div className="ai-mindmap-head">
        <span className="ai-mindmap-title">{t('aiMindMapTitle')}</span>
        <span className="ai-mindmap-count">{t('aiMindMapNodes', { count: totalNodes })}</span>
        <span className="ai-mindmap-actions">
          <button
            type="button"
            className="ai-mindmap-btn"
            onClick={fitView}
            data-tip={t('aiMindMapReset')}
            aria-label={t('aiMindMapReset')}
          >
            <svg
              viewBox="0 0 20 20"
              width="14"
              height="14"
              fill="none"
              stroke="currentColor"
              strokeWidth="1.5"
              strokeLinecap="round"
              strokeLinejoin="round"
              aria-hidden
            >
              <path d="M3.5 10a6.5 6.5 0 1 1 1.9 4.6" />
              <path d="M3.5 13.5v-3.5H7" />
            </svg>
          </button>
          <button
            type="button"
            className="ai-mindmap-btn"
            onClick={exportPng}
            data-tip={t('aiMindMapExport')}
            aria-label={t('aiMindMapExport')}
          >
            <svg
              viewBox="0 0 20 20"
              width="14"
              height="14"
              fill="none"
              stroke="currentColor"
              strokeWidth="1.5"
              strokeLinecap="round"
              strokeLinejoin="round"
              aria-hidden
            >
              <path d="M10 3v8m0 0 3-3m-3 3-3-3" />
              <path d="M4 13v2a1.5 1.5 0 0 0 1.5 1.5h9A1.5 1.5 0 0 0 16 15v-2" />
            </svg>
          </button>
        </span>
      </div>
      <div
        ref={viewportRef}
        className="ai-mindmap-viewport"
        onPointerDown={onPointerDown}
        onPointerMove={onPointerMove}
        onPointerUp={onPointerUp}
        onPointerCancel={onPointerUp}
      >
        <div
          className="ai-mindmap-canvas"
          style={{
            width: layout.width,
            height: layout.height,
            transform: `translate(${pan.x}px, ${pan.y}px) scale(${zoom})`,
          }}
        >
          <svg
            ref={svgRef}
            width={layout.width}
            height={layout.height}
            viewBox={`0 0 ${layout.width} ${layout.height}`}
            role="img"
            aria-label={t('aiMindMapTitle')}
          >
            <g className="ai-mindmap-edges">
              {layout.edges.map((edge) => (
                <path key={edge.id} d={edge.d} data-export="edge" />
              ))}
            </g>
            {layout.nodes.map((n) => {
              const isRoot = n.depth === 0
              const labelX = n.block !== undefined ? 38 : 12
              const firstY = n.lines.length > 1 ? 12 : H / 2
              return (
                <g
                  key={n.id}
                  className={`ai-mindmap-node${isRoot ? ' root' : ''}${n.collapsed ? ' collapsed' : ''}`}
                  transform={`translate(${n.x} ${n.y})`}
                  role="button"
                  tabIndex={0}
                  aria-label={n.label}
                  onClick={() => onAsk(n.label)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' || e.key === ' ') {
                      e.preventDefault()
                      onAsk(n.label)
                    }
                  }}
                >
                  <rect
                    width={W}
                    height={H}
                    rx={8}
                    data-export={isRoot ? 'root-box' : 'node-box'}
                  />
                  {n.block !== undefined && (
                    <g
                      className="ai-mindmap-badge"
                      role="button"
                      tabIndex={0}
                      aria-label={t('aiMindMapJump')}
                      data-tip={t('aiMindMapJump')}
                      onClick={(e) => {
                        e.stopPropagation()
                        onJumpBlock(n.block as number)
                      }}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' || e.key === ' ') {
                          e.preventDefault()
                          e.stopPropagation()
                          onJumpBlock(n.block as number)
                        }
                      }}
                    >
                      <rect
                        x={10}
                        y={H / 2 - 8}
                        width={22}
                        height={16}
                        rx={8}
                        data-export="badge"
                      />
                      <text
                        x={21}
                        y={H / 2}
                        textAnchor="middle"
                        dominantBaseline="central"
                        data-export="badge-text"
                      >
                        {n.block}
                      </text>
                    </g>
                  )}
                  <text
                    className="ai-mindmap-label"
                    x={labelX}
                    y={firstY}
                    dominantBaseline="central"
                    data-export={isRoot ? 'root-text' : 'node-text'}
                  >
                    {n.lines.map((line, i) => (
                      <tspan key={i} x={labelX} dy={i === 0 ? 0 : 15}>
                        {line}
                      </tspan>
                    ))}
                  </text>
                  {n.hasChildren && (
                    <g
                      className="ai-mindmap-toggle"
                      role="button"
                      tabIndex={0}
                      aria-label={n.collapsed ? t('aiMindMapExpand') : t('aiMindMapCollapse')}
                      onClick={(e) => {
                        e.stopPropagation()
                        toggleCollapse(n.id)
                      }}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' || e.key === ' ') {
                          e.preventDefault()
                          e.stopPropagation()
                          toggleCollapse(n.id)
                        }
                      }}
                    >
                      {n.collapsed ? (
                        <>
                          <rect
                            x={W - 34}
                            y={H / 2 - 8}
                            width={28}
                            height={16}
                            rx={8}
                            data-export="toggle"
                          />
                          <text
                            x={W - 20}
                            y={H / 2}
                            textAnchor="middle"
                            dominantBaseline="central"
                            data-export="toggle-text"
                          >
                            {`+${n.descendants}`}
                          </text>
                        </>
                      ) : (
                        <>
                          <circle cx={W - 15} cy={H / 2} r={7} data-export="toggle" />
                          <text
                            x={W - 15}
                            y={H / 2}
                            textAnchor="middle"
                            dominantBaseline="central"
                            data-export="toggle-text"
                          >
                            −
                          </text>
                        </>
                      )}
                    </g>
                  )}
                </g>
              )
            })}
          </svg>
        </div>
      </div>
    </div>
  )
}
