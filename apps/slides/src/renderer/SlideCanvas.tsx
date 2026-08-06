/**
 * 3.1/3.2 Konva canvas — renders one RenderSlide + selection/transform + text-edit triggering.
 */
import React, { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import {
  Stage,
  Layer,
  Rect,
  Group,
  Transformer,
  Line,
  Arrow,
  Text,
  Circle,
  Path,
} from 'react-konva'
import type Konva from 'konva'
import type {
  RenderSlide,
  RenderNode,
  ShapeRenderNode,
  TableRenderNode,
  PictureRenderNode,
  GroupRenderNode,
} from '@genoffice/pptx-render'
import { fillToKonva, isEditableText } from './konva-adapter'
import {
  computeSnap,
  computeSpacingSnap,
  type SnapTarget,
  type Guide,
  type SpacingIndicator,
} from './snap'
import { NodeBody, StaticNode } from './NodeBody'
import { useI18n } from './i18n/locale'
import {
  defaultDrawSize,
  isLineDrawKind,
  isStraightLineKind,
  resolveDrawRect,
  type DrawRect,
} from './draw-shape'
import { shapePreviewPath } from './components/gallery-previews'

/** Whether a node is a connector (read-only, no Transformer attached). */
function isConnectorNode(node: RenderNode): boolean {
  return (node.type === 'shape' || node.type === 'text') && !!(node as ShapeRenderNode).line
}

/** Find a node by sourceId (top level + recursively inside groups; while editing inside a group the selection may contain children). */
function findNodeDeep(nodes: RenderNode[], id: string): RenderNode | undefined {
  for (const n of nodes) {
    if (n.sourceId === id) return n
    if (n.type === 'group') {
      const c = findNodeDeep((n as GroupRenderNode).children, id)
      if (c) return c
    }
  }
  return undefined
}

interface Props {
  slide: RenderSlide
  selectedIds: string[]
  onSelect: (sourceId: string | null, additive?: boolean) => void
  /** caret = viewport coordinates of the double-click (editor places the caret/selects the word there; defaults to caret at end) */
  onEditText: (sourceId: string, caret?: { x: number; y: number }) => void
  /** preview=true: live preview commit during drag (not added to undo history, see EditTransformOp.preview);
   * groupId: geometry commit for a child while editing inside a group (box is in group-local coordinates) */
  onTransform: (
    sourceId: string,
    box: { x: number; y: number; w: number; h: number; rotationDeg: number },
    preview?: boolean,
    groupId?: string,
  ) => void
  /** Double-click a table cell: row/col are model coordinates */
  onEditTableCell: (sourceId: string, row: number, col: number) => void
  /** Double-click an audio/video element: trigger the playback overlay */
  onPlayMedia?: (sourceId: string) => void
  /** Right-click: hit element gives sourceId, blank area gives null; table hits include cell model coordinates */
  onContextMenu: (
    sourceId: string | null,
    clientX: number,
    clientY: number,
    cell?: { row: number; col: number },
  ) => void
  /** Drag to resize a column (col is the tc index, wPx is the column's new width in slide px) */
  onTableColResize: (sourceId: string, col: number, wPx: number) => void
  onTableRowResize?: (sourceId: string, row: number, hPx: number) => void
  images: Map<string, HTMLImageElement>
  /** Element/cell currently edited in the DOM overlay: canvas skips drawing its text (avoids ghosting under the overlay) */
  editingText?: { sourceId: string; cell?: { row: number; col: number } } | null
  /** Canvas CSS zoom (App's zoom): screen-distance semantics such as drag thresholds must be converted back to canvas coordinates */
  zoom?: number
  /** Rubber-band selection (drag a rectangle on blank area): set of elements fully inside the rectangle */
  onMarqueeSelect?: (sourceIds: string[]) => void
  /** Option+drag duplicate: original snaps back, a copy is created at the drop offset (dx/dy in slide px) */
  onDuplicateTo?: (sourceId: string, dxPx: number, dyPx: number) => void
  /** Group being edited from inside after double-click-into-group (its children are selectable/editable) */
  enteredGroupId?: string | null
  /** Double-click a group to enter in-group editing (childId is the child hit by the double-click, may be empty) */
  onEnterGroup?: (groupId: string, childId: string | null) => void
  /** Connector endpoint drag committed: new endpoints in slide px; start/end = attachment change for that end (undefined keep, null detach, object attach) */
  onEditConnectorEndpoints?: (
    sourceId: string,
    ep: {
      x1: number
      y1: number
      x2: number
      y2: number
      start?: { targetId: string; idx: number } | null
      end?: { targetId: string; idx: number } | null
    },
  ) => void
  /** Shape draw mode (ribbon gallery pick): crosshair cursor, mousedown draws over anything; click = default size */
  drawMode?: { kind: string } | null
  /** Draw gesture committed: box to insert (slide px; flipH/flipV restore a line's drag direction) */
  onDrawCommit?: (rect: DrawRect) => void
  /** Draw mode cancelled from inside the canvas (Escape) */
  onDrawCancel?: () => void
}

/**
 * Canvas bleed margin (slide px): PowerPoint's edit view draws elements that extend past the
 * slide boundary on the surrounding workspace, but pixels outside the canvas cannot be drawn.
 * Add a bleed ring around the Stage; content is wrapped in an offset Group (slide coordinate
 * system unchanged), and the Stage container is absolutely positioned at -BLEED to align.
 */
export const CANVAS_BLEED = 160

/** Walk up the Konva parent chain to the owning node_<sourceId> Group (null for background/decoration/stage). */
function nodeIdFromTarget(t: Konva.Node | null): string | null {
  while (t && t !== t.getStage()) {
    const id = typeof t.id === 'function' ? t.id() : ''
    if (id && id.startsWith('node_')) return id.slice('node_'.length)
    t = t.getParent()
  }
  return null
}

/**
 * Node count from which a slide counts as "dense" and gets its raster pressure capped:
 * every edit redraws the whole layer, and on a ~600-element page at pixelRatio 3 that keeps the
 * GPU command buffer near-full — the prime suspect for the sporadic renderer main-thread freeze.
 */
export const DENSE_SLIDE_NODE_COUNT = 300

/** Total node count including group children (a dense page may hide its elements inside groups). */
function countNodes(nodes: RenderNode[]): number {
  let n = 0
  for (const node of nodes) {
    n += 1
    if (node.type === 'group') n += countNodes((node as GroupRenderNode).children)
  }
  return n
}

/**
 * Main-canvas rasterization ratio. Zoom is a pure CSS transform on the Stage's container, so
 * without this the bitmap stays at devicePixelRatio and zoom>1 just stretches it (blurry).
 * Never below devicePixelRatio (zoom<1 keeps the old quality); capped to bound canvas memory.
 * Dense slides skip the zoom upscale and cap harder: full-layer redraw cost scales with
 * ratio² × node count, and bounding it is the cheap half of the freeze mitigation.
 */
export function canvasPixelRatio(devicePixelRatio: number, zoom: number, nodeCount = 0): number {
  if (nodeCount >= DENSE_SLIDE_NODE_COUNT) return Math.min(devicePixelRatio || 1, 2)
  return Math.min((devicePixelRatio || 1) * Math.max(zoom, 1), 3)
}

export function SlideCanvas({
  slide,
  selectedIds,
  onSelect,
  onEditText,
  onTransform,
  onEditTableCell,
  onTableColResize,
  onTableRowResize,
  onPlayMedia,
  onContextMenu,
  images,
  editingText,
  zoom = 1,
  onMarqueeSelect,
  onDuplicateTo,
  enteredGroupId,
  onEnterGroup,
  onEditConnectorEndpoints,
  drawMode,
  onDrawCommit,
  onDrawCancel,
}: Props) {
  const trRef = useRef<Konva.Transformer>(null)
  const layerRef = useRef<Konva.Layer>(null)
  const stageRef = useRef<Konva.Stage>(null)

  const nodeCount = useMemo(() => countNodes(slide.nodes), [slide])
  const dense = nodeCount >= DENSE_SLIDE_NODE_COUNT

  // zoom is committed to state only after the gesture ends (App batches wheel events),
  // so re-rasterizing here happens once per gesture, not per wheel event. The slide dep
  // covers layers (re)created between zoom changes; the resolution media query covers
  // devicePixelRatio changes (window dragged to another display) — it matches the current
  // dpr, so it must be re-armed after each change.
  useEffect(() => {
    const apply = () => {
      const ratio = canvasPixelRatio(window.devicePixelRatio, zoom, nodeCount)
      for (const l of stageRef.current?.getLayers() ?? []) {
        if (l.getCanvas().getPixelRatio() !== ratio) {
          l.getCanvas().setPixelRatio(ratio)
          l.batchDraw()
        }
      }
    }
    apply()
    let mq: MediaQueryList | null = null
    const onDprChange = () => {
      apply()
      arm()
    }
    const arm = () => {
      mq?.removeEventListener('change', onDprChange)
      mq = window.matchMedia(`(resolution: ${window.devicePixelRatio}dppx)`)
      mq.addEventListener('change', onDprChange)
    }
    arm()
    return () => mq?.removeEventListener('change', onDprChange)
  }, [zoom, slide])
  const [guides, setGuides] = useState<Guide[]>([])
  // Equal-spacing double-headed arrows (while dragging); same-size matched elements (while resizing, listed per dimension)
  const [spacing, setSpacing] = useState<SpacingIndicator[]>([])
  const [sizeMatch, setSizeMatch] = useState<{ w: string[]; h: string[] } | null>(null)
  const sizeMatchKeyRef = useRef('')
  // A marquee drag just ended on this gesture: swallow the trailing click so it doesn't select the node under the cursor
  const suppressClickRef = useRef(false)
  // Full-page background-like nodes: click-selectable but not draggable; marquee drags may start on them
  const backgroundIds = useMemo(
    () => new Set(slide.nodes.filter((n) => n.background).map((n) => n.sourceId)),
    [slide],
  )
  // Rubber-band selection rectangle (slide coordinates); null = not rubber-band selecting
  const [marquee, setMarquee] = useState<{ x1: number; y1: number; x2: number; y2: number } | null>(
    null,
  )
  const marqueeRef = useRef<{ x1: number; y1: number; x2: number; y2: number } | null>(null)

  // In-flight shape draw gesture (slide coordinates); preview only shows past the click threshold
  const drawRef = useRef<{ x1: number; y1: number; x2: number; y2: number; shift: boolean } | null>(
    null,
  )
  const [drawPreview, setDrawPreview] = useState<DrawRect | null>(null)

  // Draw mode: Escape cancels; toggling Shift mid-drag re-constrains the preview without mouse movement
  useEffect(() => {
    if (!drawMode) {
      drawRef.current = null
      setDrawPreview(null)
      return
    }
    const onKey = (ev: KeyboardEvent) => {
      if (ev.key === 'Escape') {
        drawRef.current = null
        setDrawPreview(null)
        onDrawCancel?.()
        return
      }
      if (ev.key === 'Shift') {
        const d = drawRef.current
        if (!d) return
        d.shift = ev.type === 'keydown'
        setDrawPreview(resolveDrawRect(drawMode.kind, d.x1, d.y1, d.x2, d.y2, d.shift))
      }
    }
    window.addEventListener('keydown', onKey)
    window.addEventListener('keyup', onKey)
    return () => {
      window.removeEventListener('keydown', onKey)
      window.removeEventListener('keyup', onKey)
    }
  }, [drawMode, onDrawCancel])

  // Hold Shift while rotating -> snap in 15° steps; release to restore free angle
  useEffect(() => {
    const setSnaps = (on: boolean) => {
      const tr = trRef.current
      if (!tr) return
      tr.rotationSnaps(on ? Array.from({ length: 24 }, (_, i) => i * 15) : [])
      tr.rotationSnapTolerance(7.5)
    }
    const down = (e: KeyboardEvent) => e.key === 'Shift' && setSnaps(true)
    const up = (e: KeyboardEvent) => e.key === 'Shift' && setSnaps(false)
    window.addEventListener('keydown', down)
    window.addEventListener('keyup', up)
    return () => {
      window.removeEventListener('keydown', down)
      window.removeEventListener('keyup', up)
    }
  }, [])

  // Selection changed -> attach the Transformer to all selected nodes (except connectors: selectable but no transform handles)
  useEffect(() => {
    const tr = trRef.current
    const layer = layerRef.current
    if (!tr || !layer) return
    const nodes = selectedIds
      .filter((id) => {
        const n = findNodeDeep(slide.nodes, id)
        return n && !isConnectorNode(n)
      })
      .map((id) => layer.findOne(`#node_${id}`))
      .filter((n): n is Konva.Node => !!n)
    tr.nodes(nodes)
    tr.getLayer()?.batchDraw()
  }, [selectedIds, slide, enteredGroupId])

  // Snap target edges of the other elements (excluding the dragged selection and decoration layer) + page center lines
  const snapTargets = (excludeIds: string[]): SnapTarget[] => {
    const list: SnapTarget[] = [
      { v: [slide.widthPx / 2], h: [slide.heightPx / 2] }, // Page center
      { v: [0, slide.widthPx], h: [0, slide.heightPx] }, // Page edges
    ]
    for (const n of slide.nodes) {
      if (excludeIds.includes(n.sourceId) || n.decoration || n.background) continue
      const b = n.box
      list.push({ v: [b.x, b.x + b.w / 2, b.x + b.w], h: [b.y, b.y + b.h / 2, b.y + b.h] })
    }
    return list
  }

  // Neighbor boxes for equal-spacing snapping (same exclusion rules as snapTargets)
  const spacingBoxes = (excludeIds: string[]) =>
    slide.nodes
      .filter((n) => !excludeIds.includes(n.sourceId) && !n.decoration && !n.background)
      .map((n) => ({ x: n.box.x, y: n.box.y, w: n.box.w, h: n.box.h }))

  // Bounding box of a multi-select drag (snapped as a whole; if it contains group children the coordinate systems differ, so degrade to no snapping)
  const selBBox = useMemo(() => {
    if (selectedIds.length < 2) return null
    const boxes = selectedIds
      .map((id) => slide.nodes.find((n) => n.sourceId === id))
      .filter((n): n is RenderNode => !!n)
      .map((n) => n.box)
    if (boxes.length !== selectedIds.length) return null
    const x = Math.min(...boxes.map((b) => b.x))
    const y = Math.min(...boxes.map((b) => b.y))
    return {
      x,
      y,
      w: Math.max(...boxes.map((b) => b.x + b.w)) - x,
      h: Math.max(...boxes.map((b) => b.y + b.h)) - y,
    }
  }, [selectedIds, slide])

  return (
    <Stage
      ref={stageRef}
      width={slide.widthPx + CANVAS_BLEED * 2}
      height={slide.heightPx + CANVAS_BLEED * 2}
      onMouseDown={(e) => {
        // Draw mode swallows the gesture before any selection logic (drawing works over elements too)
        if (drawMode) {
          if (e.evt.button !== 0) return
          const raw = e.target.getStage()?.getPointerPosition()
          if (!raw) return
          const x = raw.x - CANVAS_BLEED
          const y = raw.y - CANVAS_BLEED
          drawRef.current = { x1: x, y1: y, x2: x, y2: y, shift: e.evt.shiftKey }
          return
        }
        suppressClickRef.current = false
        // Blank = the Stage itself (bleed area), the slide base/background (name=slide-bg),
        // or a full-page background-like node (still click-selectable, but a drag on it
        // rubber-bands instead of moving it)
        const isBlank =
          e.target === e.target.getStage() ||
          (typeof e.target.name === 'function' && e.target.name() === 'slide-bg')
        const hitId = isBlank ? null : nodeIdFromTarget(e.target)
        const onBackground = hitId != null && backgroundIds.has(hitId)
        if (!isBlank && !onBackground) return
        if (isBlank) onSelect(null)
        // Mouse-down on blank area -> start rubber-band selection (on release, elements fully inside the rectangle are selected)
        if (e.evt.button !== 0) return
        const raw = e.target.getStage()?.getPointerPosition()
        if (!raw || !onMarqueeSelect) return
        const start = { x: raw.x - CANVAS_BLEED, y: raw.y - CANVAS_BLEED }
        marqueeRef.current = { x1: start.x, y1: start.y, x2: start.x, y2: start.y }
        setMarquee(null) // Only show after moving past the threshold
      }}
      onMouseMove={(e) => {
        if (drawMode) {
          const d = drawRef.current
          if (!d) return
          const raw = e.target.getStage()?.getPointerPosition()
          if (!raw) return
          d.x2 = raw.x - CANVAS_BLEED
          d.y2 = raw.y - CANVAS_BLEED
          d.shift = e.evt.shiftKey
          // Within 3 screen px still counts as a click; don't show the preview yet
          if (Math.hypot(d.x2 - d.x1, d.y2 - d.y1) * zoom > 3)
            setDrawPreview(resolveDrawRect(drawMode.kind, d.x1, d.y1, d.x2, d.y2, d.shift))
          return
        }
        const m = marqueeRef.current
        if (!m) return
        const raw = e.target.getStage()?.getPointerPosition()
        if (!raw) return
        m.x2 = raw.x - CANVAS_BLEED
        m.y2 = raw.y - CANVAS_BLEED
        // Within 3 screen px counts as a click; don't show the selection rectangle
        if (Math.hypot(m.x2 - m.x1, m.y2 - m.y1) * zoom > 3) setMarquee({ ...m })
      }}
      onMouseUp={() => {
        if (drawMode) {
          const d = drawRef.current
          drawRef.current = null
          setDrawPreview(null)
          if (!d || !onDrawCommit) return
          // Click (no real drag) = insert at the predefined default size, PowerPoint-style
          const rect =
            Math.hypot(d.x2 - d.x1, d.y2 - d.y1) * zoom <= 3
              ? { x: d.x1, y: d.y1, ...defaultDrawSize(drawMode.kind) }
              : resolveDrawRect(drawMode.kind, d.x1, d.y1, d.x2, d.y2, d.shift)
          onDrawCommit(rect)
          return
        }
        const m = marqueeRef.current
        marqueeRef.current = null
        if (!m) return
        setMarquee(null)
        if (Math.hypot(m.x2 - m.x1, m.y2 - m.y1) * zoom <= 3) return // A click, not a rubber-band selection
        suppressClickRef.current = true
        const [lx, rx] = m.x1 < m.x2 ? [m.x1, m.x2] : [m.x2, m.x1]
        const [ty, by] = m.y1 < m.y2 ? [m.y1, m.y2] : [m.y2, m.y1]
        const ids = slide.nodes
          .filter((n) => {
            if (n.decoration || n.background) return false
            const b = n.box
            return b.x >= lx && b.x + b.w <= rx && b.y >= ty && b.y + b.h <= by
          })
          .map((n) => n.sourceId)
        if (ids.length) onMarqueeSelect!(ids)
        else onSelect(null) // A marquee that started on a background node skipped the mousedown clear
      }}
      onContextMenu={(e) => {
        e.evt.preventDefault()
        // Walk up the parent chain to find the node_<sourceId> Group (decoration layer/background has none -> blank-area menu)
        let t: Konva.Node | null = e.target
        let sourceId: string | null = null
        while (t && t !== t.getStage()) {
          const id = typeof t.id === 'function' ? t.id() : ''
          if (id && id.startsWith('node_')) {
            sourceId = id.slice('node_'.length)
            break
          }
          t = t.getParent()
        }
        let cell: { row: number; col: number } | undefined
        if (sourceId) {
          const n = slide.nodes.find((x) => x.sourceId === sourceId)
          const raw = e.target.getStage()?.getPointerPosition()
          // Stage coordinates -> slide coordinates (remove the bleed offset)
          const pos = raw ? { x: raw.x - CANVAS_BLEED, y: raw.y - CANVAS_BLEED } : null
          if (n && n.type === 'table' && !n.box.rotationDeg && pos) {
            const hit = (n as TableRenderNode).cells.find(
              (c) =>
                pos.x - n.box.x >= c.x &&
                pos.x - n.box.x < c.x + c.w &&
                pos.y - n.box.y >= c.y &&
                pos.y - n.box.y < c.y + c.h,
            )
            if (hit) cell = { row: hit.row, col: hit.col }
          }
        }
        onContextMenu(sourceId, e.evt.clientX, e.evt.clientY, cell)
      }}
      style={{
        position: 'absolute',
        left: -CANVAS_BLEED,
        top: -CANVAS_BLEED,
        cursor: drawMode ? 'crosshair' : undefined,
      }}
    >
      {/* Draw mode disables node hit-testing: the crosshair gesture must not select/drag elements underneath */}
      <Layer ref={layerRef} x={CANVAS_BLEED} y={CANVAS_BLEED} listening={!drawMode}>
        {/* Slide base: white background + shadow (used to be the Stage's CSS background; the bleed area must show the gray workspace behind).
            Dense slides drop the decorative blur: a canvas shadow re-rasterizes on every full-layer
            redraw and is pure GPU pressure on exactly the pages where the freeze bites. */}
        <Rect
          name="slide-bg"
          x={0}
          y={0}
          width={slide.widthPx}
          height={slide.heightPx}
          fill="#ffffff"
          {...(dense
            ? { stroke: 'rgba(0,0,0,0.15)', strokeWidth: 1 }
            : { shadowColor: 'rgba(0,0,0,0.15)', shadowBlur: 16, shadowOffsetY: 2 })}
        />
        <Rect
          name="slide-bg"
          x={0}
          y={0}
          width={slide.widthPx}
          height={slide.heightPx}
          {...bgFill(slide, images)}
        />
        {slide.nodes.map((n) => (
          <NodeView
            key={n.id}
            node={n}
            onSelect={onSelect}
            onEditText={onEditText}
            onTransform={onTransform}
            onEditTableCell={onEditTableCell}
            onPlayMedia={onPlayMedia}
            onDragGuides={(g, sp) => {
              setGuides(g)
              setSpacing(sp ?? [])
            }}
            snapTargets={snapTargets}
            images={images}
            editingText={editingText}
            livePreview={selectedIds.length === 1}
            zoom={zoom}
            multiDrag={selectedIds.length > 1 && selectedIds.includes(n.sourceId)}
            onDuplicateTo={onDuplicateTo}
            entered={n.sourceId === enteredGroupId}
            onEnterGroup={onEnterGroup}
            selectedIds={selectedIds}
            selBBox={selBBox}
            spacingBoxes={spacingBoxes}
            suppressClickRef={suppressClickRef}
          />
        ))}
        {guides.map((g, i) =>
          g.axis === 'v' ? (
            <Line
              key={`v${i}`}
              points={[g.pos, 0, g.pos, slide.heightPx]}
              stroke="#ff2d55"
              strokeWidth={1}
              dash={[4, 4]}
            />
          ) : (
            <Line
              key={`h${i}`}
              points={[0, g.pos, slide.widthPx, g.pos]}
              stroke="#ff2d55"
              strokeWidth={1}
              dash={[4, 4]}
            />
          ),
        )}
        {/* Equal-spacing double-headed arrows (axis=x horizontal spacing, axis=y vertical) */}
        {spacing.map((s, i) => (
          <Arrow
            key={`sp${i}`}
            points={s.axis === 'x' ? [s.from, s.at, s.to, s.at] : [s.at, s.from, s.at, s.to]}
            stroke="#ff2d55"
            fill="#ff2d55"
            strokeWidth={1 / Math.max(zoom, 0.1)}
            pointerAtBeginning
            pointerAtEnding
            pointerLength={5 / Math.max(zoom, 0.1)}
            pointerWidth={4 / Math.max(zoom, 0.1)}
            listening={false}
          />
        ))}
        {/* Same-size match: matched elements highlighted with a dashed outline */}
        {sizeMatch &&
          [...new Set([...sizeMatch.w, ...sizeMatch.h])].map((id) => {
            const n = slide.nodes.find((x) => x.sourceId === id)
            if (!n) return null
            return (
              <Rect
                key={`sz_${id}`}
                x={n.box.x - 2}
                y={n.box.y - 2}
                width={n.box.w + 4}
                height={n.box.h + 4}
                rotation={n.box.rotationDeg}
                stroke="#ff2d55"
                strokeWidth={1 / Math.max(zoom, 0.1)}
                dash={[4, 3]}
                listening={false}
              />
            )
          })}
        {marquee && (
          <Rect
            x={Math.min(marquee.x1, marquee.x2)}
            y={Math.min(marquee.y1, marquee.y2)}
            width={Math.abs(marquee.x2 - marquee.x1)}
            height={Math.abs(marquee.y2 - marquee.y1)}
            fill="rgba(64,128,255,0.08)"
            stroke="#4080ff"
            strokeWidth={1 / Math.max(zoom, 0.1)}
            dash={[4, 4]}
            listening={false}
          />
        )}
        {/* Shape-draw preview: live ghost of the actual shape being drawn (PowerPoint drag behavior) */}
        {drawMode &&
          drawPreview &&
          (() => {
            const p = drawPreview
            // Line kinds preview with their real endpoints (flips encode leftward/upward drags)
            const sx = p.flipH ? p.x + p.w : p.x
            const sy = p.flipV ? p.y + p.h : p.y
            const ex = p.flipH ? p.x : p.x + p.w
            const ey = p.flipV ? p.y : p.y + p.h
            const lineW = 1.33 // 1pt, same as the inserted stroke
            if (isStraightLineKind(drawMode.kind)) {
              return drawMode.kind === 'line' ? (
                <Line
                  points={[sx, sy, ex, ey]}
                  stroke="#000000"
                  strokeWidth={lineW}
                  listening={false}
                />
              ) : (
                <Arrow
                  points={[sx, sy, ex, ey]}
                  stroke="#000000"
                  fill="#000000"
                  strokeWidth={lineW}
                  pointerAtEnding
                  pointerAtBeginning={drawMode.kind === 'lineArrowDouble'}
                  pointerLength={8}
                  pointerWidth={6}
                  listening={false}
                />
              )
            }
            if (isLineDrawKind(drawMode.kind)) {
              const mx = (sx + ex) / 2
              const d =
                drawMode.kind === 'lineBent'
                  ? `M ${sx} ${sy} L ${mx} ${sy} L ${mx} ${ey} L ${ex} ${ey}`
                  : `M ${sx} ${sy} C ${sx + (ex - sx) * 0.6} ${sy} ${sx + (ex - sx) * 0.4} ${ey} ${ex} ${ey}`
              return <Path data={d} stroke="#000000" strokeWidth={lineW} listening={false} />
            }
            const d = shapePreviewPath(drawMode.kind, Math.max(p.w, 1), Math.max(p.h, 1))
            if (d)
              return (
                <Path
                  x={p.x}
                  y={p.y}
                  data={d}
                  fill="rgba(196,62,28,0.45)"
                  stroke="#C43E1C"
                  strokeWidth={1 / Math.max(zoom, 0.1)}
                  listening={false}
                />
              )
            // Preset not covered by the preview geometry: dashed box fallback
            return (
              <Rect
                x={p.x}
                y={p.y}
                width={p.w}
                height={p.h}
                stroke="#4080ff"
                strokeWidth={1 / Math.max(zoom, 0.1)}
                dash={[4, 4]}
                listening={false}
              />
            )
          })()}
        {(() => {
          if (selectedIds.length !== 1) return null
          const n = slide.nodes.find((x) => x.sourceId === selectedIds[0])
          if (!n || n.type !== 'table' || n.decoration || n.box.rotationDeg) return null
          const tbl = n as TableRenderNode
          // One grip per grid boundary; only boundaries a merged cell spans across are skipped
          const spansX = (b: number) => tbl.cells.some((c) => c.x < b - 0.5 && c.x + c.w > b + 0.5)
          const spansY = (b: number) => tbl.cells.some((c) => c.y < b - 0.5 && c.y + c.h > b + 0.5)
          return [
            ...tbl.gridX.slice(1).flatMap((b, col) =>
              spansX(b)
                ? []
                : [
                    <Rect
                      key={`colgrip_${col}`}
                      x={n.box.x + b - 3}
                      y={n.box.y}
                      width={6}
                      height={n.box.h}
                      fill="transparent"
                      draggable
                      dragBoundFunc={(pos) => ({ x: pos.x, y: n.box.y + CANVAS_BLEED })}
                      onMouseEnter={(e) => {
                        const st = e.target.getStage()
                        if (st) st.container().style.cursor = 'col-resize'
                      }}
                      onMouseLeave={(e) => {
                        const st = e.target.getStage()
                        if (st) st.container().style.cursor = 'default'
                      }}
                      onDragEnd={(e) => {
                        const newW = Math.max(12, e.target.x() + 3 - (n.box.x + tbl.gridX[col]!))
                        onTableColResize(n.sourceId, col, newW)
                      }}
                    />,
                  ],
            ),
            ...(onTableRowResize
              ? tbl.gridY.slice(1).flatMap((b, row) =>
                  spansY(b)
                    ? []
                    : [
                        <Rect
                          key={`rowgrip_${row}`}
                          x={n.box.x}
                          y={n.box.y + b - 3}
                          width={n.box.w}
                          height={6}
                          fill="transparent"
                          draggable
                          dragBoundFunc={(pos) => ({ x: n.box.x + CANVAS_BLEED, y: pos.y })}
                          onMouseEnter={(e) => {
                            const st = e.target.getStage()
                            if (st) st.container().style.cursor = 'row-resize'
                          }}
                          onMouseLeave={(e) => {
                            const st = e.target.getStage()
                            if (st) st.container().style.cursor = 'default'
                          }}
                          onDragEnd={(e) => {
                            const newH = Math.max(
                              10,
                              e.target.y() + 3 - (n.box.y + tbl.gridY[row]!),
                            )
                            onTableRowResize(n.sourceId, row, newH)
                          }}
                        />,
                      ],
                )
              : []),
          ]
        })()}
        {/* Connector selection highlight: dashed bounding box (no transform handles) */}
        {selectedIds.flatMap((id) => {
          const n = slide.nodes.find((x) => x.sourceId === id)
          if (!n || !isConnectorNode(n)) return []
          const b = n.box
          return [
            <Rect
              key={`cxn_sel_${id}`}
              x={b.x - 3}
              y={b.y - 3}
              width={b.w + 6}
              height={b.h + 6}
              rotation={b.rotationDeg}
              stroke="#c43e1c"
              strokeWidth={1}
              dash={[4, 3]}
              fill="transparent"
              listening={false}
            />,
          ]
        })}
        {/* Connector endpoint handles: drag an end to reposition it; near a shape's side midpoint it snaps and attaches (stCxn/endCxn) */}
        {onEditConnectorEndpoints &&
          selectedIds.length === 1 &&
          (() => {
            const n = slide.nodes.find((x) => x.sourceId === selectedIds[0])
            if (!n || !isConnectorNode(n) || n.decoration || n.box.rotationDeg) return null
            const line = (n as ShapeRenderNode).line
            if (!line || line.points.length < 4) return null
            return (
              <ConnectorEndpointHandles
                node={n as ShapeRenderNode}
                slide={slide}
                zoom={zoom}
                onCommit={onEditConnectorEndpoints}
              />
            )
          })()}
        <Transformer
          ref={trRef}
          rotateEnabled
          // Pictures default to proportional scaling (corner handles); shapes/text boxes scale freely
          keepRatio={
            selectedIds.length > 0 &&
            selectedIds.every((id) => findNodeDeep(slide.nodes, id)?.type === 'picture')
          }
          borderStroke="#c43e1c"
          anchorStroke="#c43e1c"
          anchorFill="#ffffff"
          boundBoxFunc={(old, next) => {
            // Same-size snapping: when resizing a single, unrotated, non-picture element (keepRatio conflicts),
            // snap when width/height is close to another element's and highlight the matched element
            if (
              selectedIds.length !== 1 ||
              next.rotation !== old.rotation ||
              Math.abs(next.width - old.width) + Math.abs(next.height - old.height) < 1e-6
            )
              return next
            const sel = slide.nodes.find((n) => n.sourceId === selectedIds[0])
            if (!sel || sel.box.rotationDeg || sel.type === 'picture' || isConnectorNode(sel))
              return next
            const thr = 6 / Math.max(zoom, 0.1)
            const cands = slide.nodes.filter((n) => !n.decoration && n.sourceId !== sel.sourceId)
            const nearest = (val: number, dim: 'w' | 'h') => {
              let best: number | null = null
              for (const c of cands) {
                const v = c.box[dim]
                if (
                  Math.abs(v - val) <= thr &&
                  (best == null || Math.abs(v - val) < Math.abs(best - val))
                )
                  best = v
              }
              return best
            }
            const out = { ...next }
            const match = { w: [] as string[], h: [] as string[] }
            const w = nearest(next.width, 'w')
            if (w != null) {
              // When dragging a left-side anchor keep the right edge fixed (x compensates with width)
              if (Math.abs(next.x - old.x) > 1e-6) out.x = next.x + next.width - w
              out.width = w
              match.w = cands.filter((c) => Math.abs(c.box.w - w) < 0.5).map((c) => c.sourceId)
            }
            const h = nearest(next.height, 'h')
            if (h != null) {
              if (Math.abs(next.y - old.y) > 1e-6) out.y = next.y + next.height - h
              out.height = h
              match.h = cands.filter((c) => Math.abs(c.box.h - h) < 0.5).map((c) => c.sourceId)
            }
            const key = `${match.w.join(',')}|${match.h.join(',')}`
            if (key !== sizeMatchKeyRef.current) {
              sizeMatchKeyRef.current = key
              setSizeMatch(w != null || h != null ? match : null)
            }
            return out
          }}
          onTransformEnd={() => {
            sizeMatchKeyRef.current = ''
            setSizeMatch(null)
          }}
        />
      </Layer>
    </Stage>
  )
}

export { computeSnap }

// ── Connector endpoint handles ─────────────────────────────

interface AnchorPt {
  targetId: string
  idx: number
  x: number
  y: number
}

/**
 * Attachable connection points of the other elements: side midpoints of each box
 * (idx matches the engine's rectangle approximation — 0 top, 1 left, 2 bottom, 3 right).
 * Uses the unrotated model box, same as the engine's move-following computation.
 */
function connectorAnchors(slide: RenderSlide, excludeId: string): AnchorPt[] {
  const out: AnchorPt[] = []
  for (const n of slide.nodes) {
    if (n.decoration || n.sourceId === excludeId) continue
    if (n.type === 'placeholder-chip' || isConnectorNode(n)) continue
    const b = n.box
    out.push(
      { targetId: n.sourceId, idx: 0, x: b.x + b.w / 2, y: b.y },
      { targetId: n.sourceId, idx: 1, x: b.x, y: b.y + b.h / 2 },
      { targetId: n.sourceId, idx: 2, x: b.x + b.w / 2, y: b.y + b.h },
      { targetId: n.sourceId, idx: 3, x: b.x + b.w, y: b.y + b.h / 2 },
    )
  }
  return out
}

function ConnectorEndpointHandles({
  node,
  slide,
  zoom,
  onCommit,
}: {
  node: ShapeRenderNode
  slide: RenderSlide
  zoom: number
  onCommit: NonNullable<Props['onEditConnectorEndpoints']>
}) {
  const [drag, setDrag] = useState<{
    which: 'start' | 'end'
    x: number
    y: number
    snap: AnchorPt | null
  } | null>(null)
  const z = Math.max(zoom, 0.1)
  const pts = node.line!.points
  const start = { x: node.box.x + pts[0]!, y: node.box.y + pts[1]! }
  const end = { x: node.box.x + pts[pts.length - 2]!, y: node.box.y + pts[pts.length - 1]! }
  const anchors = useMemo(() => connectorAnchors(slide, node.sourceId), [slide, node.sourceId])
  const thr = 8 / z
  const nearestAnchor = (x: number, y: number): AnchorPt | null => {
    let best: AnchorPt | null = null
    let bestD = thr
    for (const a of anchors) {
      const d = Math.hypot(a.x - x, a.y - y)
      if (d <= bestD) {
        bestD = d
        best = a
      }
    }
    return best
  }
  const setCursor = (e: Konva.KonvaEventObject<MouseEvent>, cursor: string) => {
    const st = e.target.getStage()
    if (st) st.container().style.cursor = cursor
  }
  const handle = (which: 'start' | 'end', p: { x: number; y: number }) => (
    <Circle
      key={which}
      x={p.x}
      y={p.y}
      radius={4.5 / z}
      fill="#ffffff"
      stroke="#c43e1c"
      strokeWidth={1.5 / z}
      hitStrokeWidth={12 / z}
      draggable
      onMouseEnter={(e) => setCursor(e, 'crosshair')}
      onMouseLeave={(e) => setCursor(e, 'default')}
      onDragMove={(e) => {
        const t = e.target
        const snap = nearestAnchor(t.x(), t.y())
        if (snap) t.position({ x: snap.x, y: snap.y })
        setDrag({ which, x: t.x(), y: t.y(), snap })
      }}
      onDragEnd={(e) => {
        const t = e.target
        const x = t.x()
        const y = t.y()
        const snap = nearestAnchor(x, y)
        // Snap the Konva handle back; the model commit re-renders the connector at its new geometry
        t.position(p)
        setDrag(null)
        const att = snap ? { targetId: snap.targetId, idx: snap.idx } : null
        onCommit(
          node.sourceId,
          which === 'start'
            ? { x1: x, y1: y, x2: end.x, y2: end.y, start: att }
            : { x1: start.x, y1: start.y, x2: x, y2: y, end: att },
        )
      }}
    />
  )
  const fixed = drag?.which === 'start' ? end : start
  return (
    <>
      {/* While dragging: show all attachable anchor dots, highlight the snapped one, and preview the new run as a dashed line */}
      {drag &&
        anchors.map((a) => {
          const active = drag.snap && a.targetId === drag.snap.targetId && a.idx === drag.snap.idx
          return (
            <Circle
              key={`anchor_${a.targetId}_${a.idx}`}
              x={a.x}
              y={a.y}
              radius={(active ? 4 : 2.5) / z}
              fill={active ? '#4ea72e' : '#8e8e93'}
              listening={false}
            />
          )
        })}
      {drag && (
        <Line
          points={[fixed.x, fixed.y, drag.x, drag.y]}
          stroke="#c43e1c"
          strokeWidth={1 / z}
          dash={[4, 3]}
          listening={false}
        />
      )}
      {handle('start', start)}
      {handle('end', end)}
    </>
  )
}

function bgFill(slide: RenderSlide, images: Map<string, HTMLImageElement>) {
  const f = fillToKonva(slide.background, slide.widthPx, slide.heightPx, images)
  return Object.keys(f).length ? f : { fill: '#ffffff' }
}

interface NodeProps {
  node: RenderNode
  onSelect: (id: string | null, additive?: boolean) => void
  onEditText: (id: string, caret?: { x: number; y: number }) => void
  onTransform: Props['onTransform']
  onEditTableCell: Props['onEditTableCell']
  onPlayMedia?: Props['onPlayMedia']
  onDragGuides: (g: Guide[], spacing?: SpacingIndicator[]) => void
  snapTargets: (excludeIds: string[]) => SnapTarget[]
  /** Neighbor boxes for equal-spacing snapping (excluding the dragged selection) */
  spacingBoxes?: (excludeIds: string[]) => Array<{ x: number; y: number; w: number; h: number }>
  /** Bounding box of a multi-select drag (snapped as a whole) */
  selBBox?: { x: number; y: number; w: number; h: number } | null
  images: Map<string, HTMLImageElement>
  editingText?: Props['editingText']
  /** Live resize preview only for single selection (multi-select gestures commit per node; the preview's undo-merge semantics don't hold, so it's off) */
  livePreview: boolean
  zoom: number
  /** This node is part of a multi-select drag: disable per-node snapping (each node snapping separately would break relative positions within the selection) */
  multiDrag?: boolean
  onDuplicateTo?: Props['onDuplicateTo']
  /** This group is in in-group editing mode: children are interactive, whole-group hit area is off */
  entered?: boolean
  onEnterGroup?: Props['onEnterGroup']
  /** In-group editing: this node is a direct child of the group (geometry commits carry groupId; snapping/preview disabled) */
  insideGroupId?: string
  /** Children may double-click into text editing (only when the group is unrotated/unflipped/unscaled — the overlay can't be positioned otherwise) */
  allowChildTextEdit?: boolean
  /** Used by multiDrag computation for in-group-editing children */
  selectedIds?: string[]
  /** Set when a marquee drag just completed on this gesture: the trailing click must not select the node under the cursor */
  suppressClickRef?: React.MutableRefObject<boolean>
}

/** Throttle interval for live resize preview (ms): each preview runs an IPC round-trip + full-page relayout */
const PREVIEW_THROTTLE_MS = 50

function NodeView({
  node,
  onSelect,
  onEditText,
  onTransform,
  onEditTableCell,
  onPlayMedia,
  onDragGuides,
  snapTargets,
  spacingBoxes,
  selBBox,
  images,
  editingText,
  livePreview,
  zoom,
  multiDrag,
  onDuplicateTo,
  entered,
  onEnterGroup,
  insideGroupId,
  allowChildTextEdit,
  selectedIds,
  suppressClickRef,
}: NodeProps) {
  const { t } = useI18n()
  const { box } = node
  const groupRef = useRef<Konva.Group>(null)
  /** Most recently rendered model width/height (preview responses update it; during a gesture, scale is always relative to it) */
  const boxRef = useRef({ w: box.w, h: box.h })
  const transformingRef = useRef(false)
  const lastPreviewRef = useRef(0)

  // The Transformer's frame/scale basis defaults to getClientRect() (content bounding box). When text
  // overflows the shape box (autofit off and content too tall), overflowing glyphs inflate the bounding
  // box; after shrinking, the Transformer refresh pops the frame back to content size, which looks like
  // "resize doesn't work". Match PowerPoint: handles always sit on the shape box (model box), text may
  // overflow — override getClientRect so the Transformer measures by the model box (live boxRef value).
  useEffect(() => {
    const g = groupRef.current
    if (!g) return
    g.getClientRect = (config?: { skipTransform?: boolean; relativeTo?: Konva.Node }) => {
      const { w, h } = boxRef.current
      const local = { x: 0, y: 0, width: w, height: h }
      if (config?.skipTransform) return local
      const t = g.getAbsoluteTransform(config?.relativeTo)
      const pts = [
        t.point({ x: 0, y: 0 }),
        t.point({ x: w, y: 0 }),
        t.point({ x: w, y: h }),
        t.point({ x: 0, y: h }),
      ]
      const xs = pts.map((p) => p.x)
      const ys = pts.map((p) => p.y)
      const minX = Math.min(...xs)
      const minY = Math.min(...ys)
      return { x: minX, y: minY, width: Math.max(...xs) - minX, height: Math.max(...ys) - minY }
    }
  }, [])

  // Resize-preview response arrived (children already rendered at the new box): rebase the gesture's
  // temporary scale onto the new box so visual size (box × scale) stays continuous — frame follows the hand, content doesn't stretch or flicker.
  useLayoutEffect(() => {
    const g = groupRef.current
    const prev = boxRef.current
    if (g && transformingRef.current && (prev.w !== box.w || prev.h !== box.h)) {
      if (box.w > 0) g.scaleX(g.scaleX() * (prev.w / box.w))
      if (box.h > 0) g.scaleY(g.scaleY() * (prev.h / box.h))
    }
    boxRef.current = { w: box.w, h: box.h }
  })

  // master/layout decoration layer: read-only display, not selectable/draggable
  if (node.decoration) return <StaticNode node={node} images={images} />

  // Chips are select-only; tables/charts support p:xfrm patch persistence, so they can be dragged/resized.
  // Full-page backgrounds stay in place: a drag on them rubber-bands (Stage-level marquee) instead of moving them.
  const draggable = node.type !== 'placeholder-chip' && !node.background
  // Flip mirrors within the original box: add width offset to x then scale -1 (subtracted back when the drag reports)
  const flipOffX = box.flipH ? box.w : 0
  const flipOffY = box.flipV ? box.h : 0

  const common = {
    id: `node_${node.sourceId}`,
    x: box.x + flipOffX,
    y: box.y + flipOffY,
    rotation: box.rotationDeg,
    scaleX: box.flipH ? -1 : 1,
    scaleY: box.flipV ? -1 : 1,
    draggable,
    // A slight hand slip on click-select (3~5px is common on trackpads) shouldn't trigger a drag: Konva's
    // default 3px threshold is too sensitive, and once it becomes a drag, onDragMove snapping amplifies it into a visible 6px+ jump that commits to the model.
    // The threshold's semantics are "6 screen px": Konva compares in canvas coordinates, so divide by the canvas CSS zoom.
    dragDistance: 6 / Math.max(zoom, 0.1),
    onClick: (e: Konva.KonvaEventObject<MouseEvent>) => {
      if (suppressClickRef?.current) {
        suppressClickRef.current = false
        return
      }
      onSelect(node.sourceId, e.evt.shiftKey || e.evt.metaKey)
    },
    onTap: () => onSelect(node.sourceId),
    onDragMove: (e: Konva.KonvaEventObject<DragEvent>) => {
      // Children in in-group editing use a different coordinate system from page snap targets; don't snap
      if (insideGroupId) {
        onDragGuides([])
        return
      }
      const t = e.target
      // Snap threshold semantics are "6 screen px": convert back by the canvas CSS zoom (same as dragDistance)
      const thr = 6 / Math.max(zoom, 0.1)
      const raw = { x: t.x() - flipOffX, y: t.y() - flipOffY }
      // Multi-select drag: snap the selection bounding box as a whole. All selected
      // nodes fire their own dragmove under the Transformer's proxy drag, but each computes the same snap
      // delta from the same model bounding box + its own fixed offset -> relative positions stay unchanged.
      const bb = multiDrag
        ? selBBox && {
            x: selBBox.x + raw.x - box.x,
            y: selBBox.y + raw.y - box.y,
            w: selBBox.w,
            h: selBBox.h,
          }
        : { ...raw, w: box.w, h: box.h }
      if (!bb) {
        onDragGuides([])
        return
      }
      const exclude = multiDrag ? selectedIds! : [node.sourceId]
      const snap = computeSnap(bb, snapTargets(exclude), thr)
      let fx = snap.x
      let fy = snap.y
      // Axes not consumed by edge snapping then try equal-spacing (smart guides)
      const indicators: SpacingIndicator[] = []
      if (spacingBoxes && (fx == null || fy == null)) {
        const sp = computeSpacingSnap(
          { x: fx ?? bb.x, y: fy ?? bb.y, w: bb.w, h: bb.h },
          spacingBoxes(exclude),
          thr,
        )
        if (fx == null && sp.x != null) {
          fx = sp.x
          indicators.push(...sp.indicators.filter((i) => i.axis === 'x'))
        }
        if (fy == null && sp.y != null) {
          fy = sp.y
          indicators.push(...sp.indicators.filter((i) => i.axis === 'y'))
        }
      }
      if (fx != null) t.x(raw.x + (fx - bb.x) + flipOffX)
      if (fy != null) t.y(raw.y + (fy - bb.y) + flipOffY)
      onDragGuides(snap.guides, indicators)
    },
    onDragEnd: (e: Konva.KonvaEventObject<DragEvent>) => {
      onDragGuides([])
      const dropX = e.target.x() - flipOffX
      const dropY = e.target.y() - flipOffY
      // Option+drag = duplicate at the drop point;
      // original snaps back (model unchanged, the Konva node position must be manually restored); multi-select/in-group gestures not supported yet
      if (e.evt?.altKey && onDuplicateTo && !multiDrag && !insideGroupId) {
        e.target.position({ x: box.x + flipOffX, y: box.y + flipOffY })
        onDuplicateTo(node.sourceId, dropX - box.x, dropY - box.y)
        return
      }
      onTransform(
        node.sourceId,
        {
          x: dropX,
          y: dropY,
          w: box.w,
          h: box.h,
          rotationDeg: box.rotationDeg,
        },
        undefined,
        insideGroupId,
      )
    },
    onTransformStart: () => {
      transformingRef.current = true
      lastPreviewRef.current = 0
    },
    onTransform: (e: Konva.KonvaEventObject<Event>) => {
      // Live preview during drag: re-lay out text at the new box size (font size unchanged).
      // Commit width/height only (x/y are managed by the Konva gesture and land in the model with the final commit on release).
      // Offset conversion for flipped elements relies on a one-shot back-calculation at gesture end; preview stays disabled for them.
      if (!livePreview || box.flipH || box.flipV) return
      const t = e.target
      const sx = Math.abs(t.scaleX())
      const sy = Math.abs(t.scaleY())
      if (Math.abs(sx - 1) < 1e-3 && Math.abs(sy - 1) < 1e-3) return // Pure rotation / not scaled
      const now = performance.now()
      if (now - lastPreviewRef.current < PREVIEW_THROTTLE_MS) return
      lastPreviewRef.current = now
      onTransform(
        node.sourceId,
        {
          x: box.x,
          y: box.y,
          w: boxRef.current.w * sx,
          h: boxRef.current.h * sy,
          rotationDeg: box.rotationDeg,
        },
        true,
      )
    },
    onTransformEnd: (e: Konva.KonvaEventObject<Event>) => {
      transformingRef.current = false
      const t = e.target
      const sx = Math.abs(t.scaleX())
      const sy = Math.abs(t.scaleY())
      onTransform(
        node.sourceId,
        {
          x: t.x() - flipOffX * sx,
          y: t.y() - flipOffY * sy,
          w: box.w * sx,
          h: box.h * sy,
          rotationDeg: t.rotation(),
        },
        undefined,
        insideGroupId,
      )
      t.scaleX(t.scaleX() < 0 ? -1 : 1)
      t.scaleY(t.scaleY() < 0 ? -1 : 1)
    },
  }

  // Group in in-group editing mode: whole-group hit area off, children become interactive NodeViews (geometry commits carry groupId)
  if (node.type === 'group' && entered) {
    const g = node as GroupRenderNode
    // The DOM text overlay doesn't follow group rotation/flip; in those cases children can't double-click into text editing.
    // ext/chExt scaling is already baked into child geometry (including text layout), so it no longer affects overlay alignment.
    const plain = !box.rotationDeg && !box.flipH && !box.flipV
    return (
      <Group
        id={`node_${node.sourceId}`}
        x={box.x + flipOffX}
        y={box.y + flipOffY}
        rotation={box.rotationDeg}
        scaleX={box.flipH ? -1 : 1}
        scaleY={box.flipV ? -1 : 1}
      >
        {/* In-group boundary indicator (dashed frame while editing a group) */}
        <Rect
          width={box.w}
          height={box.h}
          stroke="#8e8e93"
          strokeWidth={1 / Math.max(zoom, 0.1)}
          dash={[3, 3]}
          listening={false}
        />
        <Group>
          {g.children.map((c) => (
            <NodeView
              key={c.id}
              node={c}
              onSelect={onSelect}
              onEditText={onEditText}
              onTransform={onTransform}
              onEditTableCell={onEditTableCell}
              onPlayMedia={onPlayMedia}
              onDragGuides={onDragGuides}
              snapTargets={snapTargets}
              images={images}
              editingText={editingText}
              livePreview={false}
              zoom={zoom}
              multiDrag={(selectedIds?.length ?? 0) > 1 && !!selectedIds?.includes(c.sourceId)}
              insideGroupId={node.sourceId}
              allowChildTextEdit={plain}
              suppressClickRef={suppressClickRef}
            />
          ))}
        </Group>
      </Group>
    )
  }

  const editable = isEditableText(node) && (!insideGroupId || allowChildTextEdit)
  // Double-click a group = enter in-group editing and select the child hit by the double-click (pointer converted to group-local coordinates, bounding-box hit)
  const onGroupDblClick = (e: Konva.KonvaEventObject<Event>) => {
    if (!onEnterGroup) return
    const g = node as GroupRenderNode
    const p = e.target.getStage()?.getPointerPosition()
    let childId: string | null = null
    if (p && groupRef.current) {
      // children box already includes ext/chExt scaling (group-local px); compare directly after converting the pointer
      const local = groupRef.current.getAbsoluteTransform().copy().invert().point(p)
      const lx = local.x
      const ly = local.y
      const hit = [...g.children]
        .reverse()
        .find(
          (c) =>
            lx >= c.box.x && lx <= c.box.x + c.box.w && ly >= c.box.y && ly <= c.box.y + c.box.h,
        )
      childId = hit?.sourceId ?? null
    }
    onEnterGroup(node.sourceId, childId)
  }
  // Table: double-click hits a cell (pointer coordinates -> table-local coordinates; editing rotated tables not supported yet)
  const onTableDblClick = (e: Konva.KonvaEventObject<Event>) => {
    if (node.type !== 'table' || box.rotationDeg) return
    const pos = e.target.getStage()?.getPointerPosition()
    if (!pos) return
    const rx = pos.x - CANVAS_BLEED - box.x
    const ry = pos.y - CANVAS_BLEED - box.y
    const cell = (node as TableRenderNode).cells.find(
      (c) => rx >= c.x && rx < c.x + c.w && ry >= c.y && ry < c.y + c.h,
    )
    if (cell) onEditTableCell(node.sourceId, cell.row, cell.col)
  }
  // Audio/video (image is the poster frame): double-click opens the playback overlay
  const isMedia = node.type === 'picture' && !!(node as PictureRenderNode).media && !!onPlayMedia
  // Empty placeholder: canvas draws a gray click hint (edit canvas only, not in thumbnails/export)
  const phPrompt = (() => {
    if (node.type !== 'shape' && node.type !== 'text') return null
    const sh = node as ShapeRenderNode
    const kind = sh.placeholder
    if (!kind || !['title', 'ctrTitle', 'subTitle', 'body'].includes(kind)) return null
    if (sh.text?.lines.some((l) => l.runs.some((r) => !r.isBullet && r.text.trim()))) return null
    if (editingText && editingText.sourceId === node.sourceId) return null
    return t(
      kind === 'subTitle'
        ? 'appPhPromptSubtitle'
        : kind === 'body'
          ? 'appPhPromptBody'
          : 'appPhPromptTitle',
    )
  })()
  return (
    <Group
      ref={groupRef}
      {...common}
      {...(editable
        ? {
            onDblClick: (e: Konva.KonvaEventObject<MouseEvent>) =>
              onEditText(node.sourceId, { x: e.evt.clientX, y: e.evt.clientY }),
            onDblTap: () => onEditText(node.sourceId),
          }
        : node.type === 'group' && !insideGroupId && onEnterGroup
          ? { onDblClick: onGroupDblClick, onDblTap: onGroupDblClick }
          : node.type === 'table' && !insideGroupId
            ? { onDblClick: onTableDblClick, onDblTap: onTableDblClick }
            : isMedia && !insideGroupId
              ? {
                  onDblClick: () => onPlayMedia!(node.sourceId),
                  onDblTap: () => onPlayMedia!(node.sourceId),
                }
              : {})}
    >
      {/* group children don't take hits (listening=false); add a transparent hit area so the whole group can be selected/dragged */}
      {node.type === 'group' && <Rect width={box.w} height={box.h} fill="transparent" />}
      <NodeBody
        node={node}
        images={images}
        hideText={!!editingText && editingText.sourceId === node.sourceId && !editingText.cell}
        hideCellText={
          editingText && editingText.sourceId === node.sourceId ? editingText.cell : undefined
        }
      />
      {/* Multi-select: PowerPoint-style per-element border so every selected element is visibly selected
          (the shared Transformer only draws one combined box). Lives inside the node group so it follows drags/transforms. */}
      {multiDrag && !isConnectorNode(node) && (
        <Rect
          width={box.w}
          height={box.h}
          stroke="#c43e1c"
          strokeWidth={1 / Math.max(zoom, 0.1)}
          strokeScaleEnabled={false}
          listening={false}
        />
      )}
      {phPrompt &&
        (() => {
          const sh = node as ShapeRenderNode
          const ins = sh.text?.insets ?? { l: 8, t: 4, r: 8, b: 4 }
          return (
            <Text
              text={phPrompt}
              x={ins.l}
              y={ins.t}
              width={Math.max(box.w - ins.l - ins.r, 20)}
              height={Math.max(box.h - ins.t - ins.b, 16)}
              align={sh.placeholder === 'body' ? 'left' : 'center'}
              verticalAlign={sh.text?.anchor ?? (sh.placeholder === 'body' ? 'top' : 'middle')}
              fontSize={Math.min(28, Math.max(14, box.h * 0.22))}
              fill="#8e8e93"
            />
          )
        })()}
    </Group>
  )
}

// Vertical/horizontal alignment offsets are already baked into glyph coordinates by pptx-render's layoutText; the renderer draws what it gets.
// Fallback reference (the Transformer frame still needs ShapeRenderNode type semantics)
export type { ShapeRenderNode }
