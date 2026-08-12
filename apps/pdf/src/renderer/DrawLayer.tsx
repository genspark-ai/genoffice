import { useRef, useState } from 'react'
import type { PointerEvent as ReactPointerEvent, ReactElement } from 'react'
import { pdfToView, viewToPdf } from './annotations'
import type { PageGeom } from './annotations'
import type { DrawingInput } from '../shared/ipc'

export type DrawTool = 'ink' | 'rect' | 'ellipse' | 'line' | 'arrow' | 'note'

/** Displayed-pixel box (scaled) */
interface Box {
  x: number
  y: number
  w: number
  h: number
}

export interface LocalDrawing {
  id: string
  input: DrawingInput
  /** Present when this drawing is the visual value placed into an AcroForm /Sig widget. */
  formWidgetId?: string
}

/** 5-color palette (0-1 rgb), same visual language as the markup floating bar */
export const DRAW_COLORS: { name: string; rgb: [number, number, number] }[] = [
  { name: 'red', rgb: [0.86, 0.22, 0.18] },
  { name: 'yellow', rgb: [1, 0.78, 0.13] },
  { name: 'green', rgb: [0.13, 0.65, 0.35] },
  { name: 'blue', rgb: [0.17, 0.4, 1] },
  { name: 'black', rgb: [0.15, 0.15, 0.15] },
]

export const cssRgb = (c: readonly [number, number, number]): string =>
  `rgb(${Math.round(c[0] * 255)}, ${Math.round(c[1] * 255)}, ${Math.round(c[2] * 255)})`

/** Drawing SVG paths: PDF space → displayed page pixels */
function toView(geom: PageGeom, scale: number, x: number, y: number): [number, number] {
  const [vx, vy] = pdfToView(geom, x, y)
  return [vx * scale, vy * scale]
}

function drawingShape(
  d: DrawingInput,
  geom: PageGeom,
  scale: number,
  /** Invisible fat outline: rendered under pointer-events:stroke so thin shapes are grabbable */
  hit = false,
): ReactElement | null {
  if (d.kind === 'note') return null
  if (d.kind === 'image') {
    if (hit) return null // the image body is already a full-area hit target
    const [ax, ay] = toView(geom, scale, d.rect[0], d.rect[1])
    const [bx, by] = toView(geom, scale, d.rect[2], d.rect[3])
    return (
      <image
        href={`data:image/png;base64,${d.image}`}
        x={Math.min(ax, bx)}
        y={Math.min(ay, by)}
        width={Math.abs(bx - ax)}
        height={Math.abs(by - ay)}
        preserveAspectRatio="none"
      />
    )
  }
  const stroke = hit ? 'transparent' : cssRgb(d.color)
  const w = hit ? Math.max(16, d.width * scale * 4) : d.width * scale
  if (d.kind === 'ink') {
    return (
      <>
        {d.paths.map((path, i) => {
          const pts: string[] = []
          for (let j = 0; j < path.length; j += 2) {
            const [vx, vy] = toView(geom, scale, path[j]!, path[j + 1]!)
            pts.push(`${vx},${vy}`)
          }
          return (
            <polyline
              key={i}
              points={pts.join(' ')}
              fill="none"
              stroke={stroke}
              strokeWidth={w}
              strokeLinecap="round"
              strokeLinejoin="round"
            />
          )
        })}
      </>
    )
  }
  if (d.kind === 'rect' || d.kind === 'ellipse') {
    const [ax, ay] = toView(geom, scale, d.rect[0], d.rect[1])
    const [bx, by] = toView(geom, scale, d.rect[2], d.rect[3])
    const [x, y] = [Math.min(ax, bx), Math.min(ay, by)]
    const [rw, rh] = [Math.abs(bx - ax), Math.abs(by - ay)]
    return d.kind === 'rect' ? (
      <rect x={x} y={y} width={rw} height={rh} fill="none" stroke={stroke} strokeWidth={w} />
    ) : (
      <ellipse
        cx={x + rw / 2}
        cy={y + rh / 2}
        rx={rw / 2}
        ry={rh / 2}
        fill="none"
        stroke={stroke}
        strokeWidth={w}
      />
    )
  }
  const [fx, fy] = toView(geom, scale, d.from[0], d.from[1])
  const [tx, ty] = toView(geom, scale, d.to[0], d.to[1])
  const head: ReactElement[] = []
  if (d.kind === 'arrow') {
    const ang = Math.atan2(ty - fy, tx - fx)
    const len = Math.max(9 * scale, w * 4.5)
    for (const [i, off] of [-0.45, 0.45].entries()) {
      head.push(
        <line
          key={i}
          x1={tx}
          y1={ty}
          x2={tx - len * Math.cos(ang + off)}
          y2={ty - len * Math.sin(ang + off)}
          stroke={stroke}
          strokeWidth={w}
          strokeLinecap="round"
        />,
      )
    }
  }
  return (
    <>
      <line x1={fx} y1={fy} x2={tx} y2={ty} stroke={stroke} strokeWidth={w} strokeLinecap="round" />
      {head}
    </>
  )
}

/**
 * Draw layer: with a tool active it takes over pointer events to draw on the page;
 * otherwise it only renders existing shapes (click to select, drag to move; deletion is
 * an explicit App action). Notes are placed at the click point; content is entered via
 * an App dialog.
 */
export function DrawLayer({
  geom,
  scale,
  pageWidth,
  pageHeight,
  drawings,
  tool,
  color,
  strokeWidth,
  selectedId,
  selectTitle,
  onCommit,
  onNoteAt,
  onSelect,
  onMove,
  onResize,
}: {
  geom: PageGeom
  scale: number
  pageWidth: number
  pageHeight: number
  drawings: LocalDrawing[]
  tool: DrawTool | null
  color: [number, number, number]
  strokeWidth: number
  selectedId: string | null
  selectTitle: string
  onCommit: (input: DrawingInput) => void
  onNoteAt: (at: [number, number]) => void
  onSelect: (id: string, x: number, y: number) => void
  /** Move a shape by a PDF-space delta; omit to disable dragging (read-only) */
  onMove?: (id: string, dx: number, dy: number) => void
  /** Replace an image drawing's PDF-space rect (corner-handle resize); omit to disable */
  onResize?: (id: string, rect: [number, number, number, number]) => void
}): ReactElement {
  // In-progress stroke/drag, all in PDF-space coordinates
  const [live, setLive] = useState<DrawingInput | null>(null)
  const startRef = useRef<[number, number] | null>(null)
  const inkRef = useRef<number[]>([])
  // Dragging an existing shape: view coords (page-relative, scale=1)
  const [drag, setDrag] = useState<{
    id: string
    from: [number, number]
    to: [number, number]
  } | null>(null)
  // Corner-resizing a selected image: boxes in displayed px (scaled)
  const [resize, setResize] = useState<{ id: string; corner: number; start: Box; box: Box } | null>(
    null,
  )

  const toPdf = (e: ReactPointerEvent): [number, number] => {
    const box = e.currentTarget.getBoundingClientRect()
    return viewToPdf(geom, (e.clientX - box.left) / scale, (e.clientY - box.top) / scale)
  }

  const onPointerDown = (e: ReactPointerEvent) => {
    if (!tool || e.button !== 0) return
    e.preventDefault()
    const at = toPdf(e)
    if (tool === 'note') {
      onNoteAt(at)
      return
    }
    e.currentTarget.setPointerCapture(e.pointerId)
    startRef.current = at
    if (tool === 'ink') {
      inkRef.current = [at[0], at[1]]
      setLive({ kind: 'ink', pageIndex: 0, color, width: strokeWidth, paths: [inkRef.current] })
    }
  }

  const onPointerMove = (e: ReactPointerEvent) => {
    if (!tool || !startRef.current) return
    const at = toPdf(e)
    if (tool === 'ink') {
      inkRef.current = [...inkRef.current, at[0], at[1]]
      setLive({ kind: 'ink', pageIndex: 0, color, width: strokeWidth, paths: [inkRef.current] })
    } else if (tool === 'rect' || tool === 'ellipse') {
      const [sx, sy] = startRef.current
      setLive({
        kind: tool,
        pageIndex: 0,
        color,
        width: strokeWidth,
        rect: [Math.min(sx, at[0]), Math.min(sy, at[1]), Math.max(sx, at[0]), Math.max(sy, at[1])],
      })
    } else if (tool === 'line' || tool === 'arrow') {
      setLive({
        kind: tool,
        pageIndex: 0,
        color,
        width: strokeWidth,
        from: startRef.current,
        to: at,
      })
    }
  }

  const onPointerUp = () => {
    const pending = live
    startRef.current = null
    inkRef.current = []
    setLive(null)
    if (!pending) return
    // Discard empty shapes from a click-and-release
    if (pending.kind === 'ink' && (pending.paths[0]?.length ?? 0) < 6) return
    if (pending.kind === 'rect' || pending.kind === 'ellipse') {
      const [x1, y1, x2, y2] = pending.rect
      if (x2 - x1 < 3 && y2 - y1 < 3) return
    }
    if (pending.kind === 'line' || pending.kind === 'arrow') {
      if (Math.hypot(pending.to[0] - pending.from[0], pending.to[1] - pending.from[1]) < 4) return
    }
    onCommit(pending)
  }

  // ── Drag-to-move an existing shape (no tool active) ──

  const shapeViewPos = (e: ReactPointerEvent<SVGGElement>): [number, number] => {
    const box = e.currentTarget.ownerSVGElement!.getBoundingClientRect()
    return [(e.clientX - box.left) / scale, (e.clientY - box.top) / scale]
  }

  const shapeDown = (e: ReactPointerEvent<SVGGElement>, id: string) => {
    if (tool || e.button !== 0) return
    // Without preventDefault the drag would also start a text selection in the
    // text layer underneath (and pop the markup bar on release)
    e.preventDefault()
    e.stopPropagation()
    e.currentTarget.setPointerCapture(e.pointerId)
    const at = shapeViewPos(e)
    setDrag({ id, from: at, to: at })
  }

  const shapeMove = (e: ReactPointerEvent<SVGGElement>) => {
    if (drag && onMove) setDrag({ ...drag, to: shapeViewPos(e) })
  }

  const shapeUp = (e: ReactPointerEvent<SVGGElement>) => {
    const d = drag
    setDrag(null)
    if (!d) return
    // A sub-3px drag is a click: select instead of committing a no-op move
    if (!onMove || Math.hypot(d.to[0] - d.from[0], d.to[1] - d.from[1]) * scale < 3) {
      onSelect(d.id, e.clientX, e.clientY)
      return
    }
    const [ax, ay] = viewToPdf(geom, d.from[0], d.from[1])
    const [bx, by] = viewToPdf(geom, d.to[0], d.to[1])
    onMove(d.id, bx - ax, by - ay)
  }

  // ── Corner-handle resize for a selected image (aspect ratio preserved) ──

  const imgBox = (d: LocalDrawing): Box | null => {
    if (d.input.kind !== 'image') return null
    const [ax, ay] = toView(geom, scale, d.input.rect[0], d.input.rect[1])
    const [bx, by] = toView(geom, scale, d.input.rect[2], d.input.rect[3])
    return { x: Math.min(ax, bx), y: Math.min(ay, by), w: Math.abs(bx - ax), h: Math.abs(by - ay) }
  }

  const handleDown = (
    e: ReactPointerEvent<SVGCircleElement>,
    id: string,
    corner: number,
    box: Box,
  ) => {
    if (e.button !== 0 || !onResize) return
    e.preventDefault()
    e.stopPropagation()
    e.currentTarget.setPointerCapture(e.pointerId)
    setResize({ id, corner, start: box, box })
  }

  const handleMove = (e: ReactPointerEvent<SVGCircleElement>) => {
    if (!resize) return
    const svgBox = e.currentTarget.ownerSVGElement!.getBoundingClientRect()
    const px = e.clientX - svgBox.left
    const py = e.clientY - svgBox.top
    const { start, corner } = resize
    // Fixed point is the opposite corner; the scale factor follows the dominant axis
    const fx = corner % 2 === 0 ? start.x + start.w : start.x
    const fy = corner < 2 ? start.y + start.h : start.y
    const k = Math.max(
      Math.abs(px - fx) / start.w,
      Math.abs(py - fy) / start.h,
      16 / Math.max(start.w, start.h),
    )
    const w = start.w * k
    const h = start.h * k
    setResize({
      ...resize,
      box: { x: corner % 2 === 0 ? fx - w : fx, y: corner < 2 ? fy - h : fy, w, h },
    })
  }

  const handleUp = () => {
    const r = resize
    setResize(null)
    if (!r || !onResize) return
    if (Math.abs(r.box.w - r.start.w) < 2 && Math.abs(r.box.h - r.start.h) < 2) return
    const [ax, ay] = viewToPdf(geom, r.box.x / scale, r.box.y / scale)
    const [bx, by] = viewToPdf(geom, (r.box.x + r.box.w) / scale, (r.box.y + r.box.h) / scale)
    onResize(r.id, [Math.min(ax, bx), Math.min(ay, by), Math.max(ax, bx), Math.max(ay, by)])
  }

  /** Live transform mapping a shape's committed box onto the in-progress resize box */
  const resizeTransform = (r: NonNullable<typeof resize>): string => {
    const k = r.box.w / r.start.w
    return `translate(${r.box.x - k * r.start.x} ${r.box.y - k * r.start.y}) scale(${k})`
  }

  return (
    <>
      <svg
        className={`pdf-draw-layer${tool ? ' pdf-draw-active' : ''}`}
        width={pageWidth * scale}
        height={pageHeight * scale}
        onPointerDown={onPointerDown}
        onPointerMove={onPointerMove}
        onPointerUp={onPointerUp}
        onPointerCancel={onPointerUp}
      >
        {drawings.map((d) =>
          d.input.kind === 'note' ? null : (
            <g
              key={d.id}
              className={`pdf-draw-shape${d.id === selectedId ? ' pdf-draw-selected' : ''}`}
              transform={
                drag?.id === d.id
                  ? `translate(${(drag.to[0] - drag.from[0]) * scale} ${(drag.to[1] - drag.from[1]) * scale})`
                  : resize?.id === d.id
                    ? resizeTransform(resize)
                    : undefined
              }
              onPointerDown={(e) => shapeDown(e, d.id)}
              onPointerMove={shapeMove}
              onPointerUp={shapeUp}
              onPointerCancel={() => setDrag(null)}
              style={{
                pointerEvents: tool ? 'none' : d.input.kind === 'image' ? 'auto' : 'stroke',
              }}
            >
              <title>{selectTitle}</title>
              {!tool && drawingShape(d.input, geom, scale, true)}
              {drawingShape(d.input, geom, scale)}
            </g>
          ),
        )}
        {/* Corner handles: shown for a selected image, drag to scale (aspect kept) */}
        {!tool &&
          !drag &&
          onResize &&
          selectedId &&
          (() => {
            const sel = drawings.find((d) => d.id === selectedId)
            const committed = sel ? imgBox(sel) : null
            if (!sel || !committed) return null
            const box = resize?.id === sel.id ? resize.box : committed
            const corners: [number, number, string][] = [
              [box.x, box.y, 'nwse-resize'],
              [box.x + box.w, box.y, 'nesw-resize'],
              [box.x, box.y + box.h, 'nesw-resize'],
              [box.x + box.w, box.y + box.h, 'nwse-resize'],
            ]
            return corners.map(([cx, cy, cursor], i) => (
              <circle
                key={i}
                className="pdf-draw-handle"
                cx={cx}
                cy={cy}
                r={5}
                style={{ cursor, pointerEvents: 'all' }}
                onPointerDown={(e) => handleDown(e, sel.id, i, committed)}
                onPointerMove={handleMove}
                onPointerUp={handleUp}
                onPointerCancel={() => setResize(null)}
              />
            ))
          })()}
        {live && drawingShape(live, geom, scale)}
      </svg>
      {drawings
        .filter((d) => d.input.kind === 'note')
        .map((d) => {
          const note = d.input as Extract<DrawingInput, { kind: 'note' }>
          const [vx, vy] = toView(geom, scale, note.at[0], note.at[1])
          return (
            <button
              key={d.id}
              className={`pdf-note-pin${d.id === selectedId ? ' pdf-note-pin-selected' : ''}`}
              style={{ left: vx, top: vy - 20, background: cssRgb(note.color) }}
              data-tip={`${note.contents}\n\n${selectTitle}`}
              aria-label={`${note.contents}\n\n${selectTitle}`}
              onClick={(e) => onSelect(d.id, e.clientX, e.clientY)}
            >
              <svg
                width="11"
                height="11"
                viewBox="0 0 16 16"
                fill="none"
                stroke="#fff"
                strokeWidth="1.6"
                aria-hidden
              >
                <path d="M2.5 3.5h11v8h-6l-3 2.5V11.5h-2z" strokeLinejoin="round" />
              </svg>
            </button>
          )
        })}
    </>
  )
}
