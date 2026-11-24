import type { DrawingInput } from '../shared/ipc'

export type ShapeAnnotType = 'ink' | 'square' | 'circle' | 'line' | 'polygon' | 'caret'

/** pdf.js annotationType → the shape subtypes the Draw tools write */
export const SHAPE_TYPE_BY_ANNOT: Record<number, ShapeAnnotType> = {
  4: 'line',
  5: 'square',
  6: 'circle',
  7: 'polygon',
  14: 'caret',
  15: 'ink',
}

type Rect = [number, number, number, number]

/** What a pointer can hit, in PDF user space: stroked polylines, plus a filled box
    for shapes with no usable geometry */
export interface ShapeOutline {
  paths: number[][]
  box?: Rect
  /** Extra reach around the paths: half the stroke width, or the band a border may lie in */
  slack?: number
}

/** A drawn shape annotation already saved in the file */
export interface SavedShapeAnnot {
  pageIndex: number
  objNum: number
  type: ShapeAnnotType
  rect: Rect
  outline: ShapeOutline
}

const rectPath = ([x1, y1, x2, y2]: Rect): number[] => [x1, y1, x2, y1, x2, y2, x1, y2, x1, y1]

const ELLIPSE_STEPS = 48

function ellipsePath([x1, y1, x2, y2]: Rect): number[] {
  const [cx, cy, rx, ry] = [(x1 + x2) / 2, (y1 + y2) / 2, (x2 - x1) / 2, (y2 - y1) / 2]
  const path: number[] = []
  for (let i = 0; i <= ELLIPSE_STEPS; i++) {
    const a = (i / ELLIPSE_STEPS) * 2 * Math.PI
    path.push(cx + rx * Math.cos(a), cy + ry * Math.sin(a))
  }
  return path
}

const CARET_HALF = 8

/** Outline of a pending drawing; null for kinds the eraser leaves alone */
export function drawingOutline(d: DrawingInput): ShapeOutline | null {
  switch (d.kind) {
    case 'ink':
      return { paths: d.paths, slack: d.width / 2 }
    case 'rect':
      return { paths: [rectPath(d.rect)], slack: d.width / 2 }
    case 'ellipse':
      return { paths: [ellipsePath(d.rect)], slack: d.width / 2 }
    case 'line':
    case 'arrow':
    case 'measure':
      return { paths: [[...d.from, ...d.to]], slack: d.width / 2 }
    case 'polygon':
      return { paths: [[...d.vertices.flat(), ...d.vertices[0]!]], slack: d.width / 2 }
    case 'caret': {
      const [x, y] = d.point
      return { paths: [], box: [x - CARET_HALF, y - CARET_HALF, x + CARET_HALF, y + CARET_HALF] }
    }
    default:
      return null
  }
}

/** The fields of a pdf.js annotation the shape tools read */
export interface PdfJsShapeData {
  id: string
  annotationType: number
  rect: number[]
  inkLists?: ArrayLike<number>[] | null
  vertices?: ArrayLike<number> | null
  lineCoordinates?: number[] | null
  borderStyle?: { width?: number } | null
}

// save-pdf.ts pads /Rect by width + 2 around the stroke; other writers put the stroke
// on the rect, half a width in. The band from the rect to our stroke's inner edge covers both.
function borderBand(rect: Rect, width: number): { rect: Rect; slack: number } {
  const [x1, y1, x2, y2] = [
    Math.min(rect[0], rect[2]),
    Math.min(rect[1], rect[3]),
    Math.max(rect[0], rect[2]),
    Math.max(rect[1], rect[3]),
  ]
  const half = Math.min((1.5 * width + 2) / 2, (x2 - x1) / 2, (y2 - y1) / 2)
  return { rect: [x1 + half, y1 + half, x2 - half, y2 - half], slack: half }
}

export function toSavedShape(a: PdfJsShapeData, pageIndex: number): SavedShapeAnnot | null {
  const type = SHAPE_TYPE_BY_ANNOT[a.annotationType]
  // Only ref-backed annots can be addressed for deletion (id "123R" → object 123)
  const objNum = /^(\d+)R$/.exec(a.id)
  if (!type || !objNum || a.rect.length < 4) return null
  const rect: Rect = [a.rect[0]!, a.rect[1]!, a.rect[2]!, a.rect[3]!]
  const width = Math.max(0, a.borderStyle?.width ?? 1)
  const slack = width / 2
  let outline: ShapeOutline = { paths: [], box: rect }
  if (type === 'ink' && a.inkLists?.length) {
    outline = { paths: a.inkLists.map((list) => Array.from(list)), slack }
  } else if (type === 'polygon' && a.vertices && a.vertices.length >= 6) {
    const v = Array.from(a.vertices)
    outline = { paths: [[...v, v[0]!, v[1]!]], slack }
  } else if (type === 'line' && a.lineCoordinates?.length === 4) {
    // pdf.js normalizes /L, which loses the line's direction: either diagonal may be it
    const [x1, y1, x2, y2] = a.lineCoordinates as Rect
    outline = {
      paths: [
        [x1, y1, x2, y2],
        [x1, y2, x2, y1],
      ],
      slack,
    }
  } else if (type === 'square' || type === 'circle') {
    const band = borderBand(rect, width)
    outline = {
      paths: [type === 'square' ? rectPath(band.rect) : ellipsePath(band.rect)],
      slack: band.slack,
    }
  }
  return { pageIndex, objNum: Number(objNum[1]), type, rect, outline }
}

function segmentDistance(
  px: number,
  py: number,
  ax: number,
  ay: number,
  bx: number,
  by: number,
): number {
  const [dx, dy] = [bx - ax, by - ay]
  const len2 = dx * dx + dy * dy
  const t = len2 === 0 ? 0 : Math.max(0, Math.min(1, ((px - ax) * dx + (py - ay) * dy) / len2))
  return Math.hypot(px - (ax + t * dx), py - (ay + t * dy))
}

/** Whether (x, y) lies within `tol` of the outline's strokes, or inside its box */
export function outlineHit(o: ShapeOutline, x: number, y: number, tol: number): boolean {
  if (o.box) {
    const [x1, y1, x2, y2] = o.box
    if (
      x >= Math.min(x1, x2) - tol &&
      x <= Math.max(x1, x2) + tol &&
      y >= Math.min(y1, y2) - tol &&
      y <= Math.max(y1, y2) + tol
    )
      return true
  }
  const reach = tol + (o.slack ?? 0)
  return o.paths.some((p) => {
    if (p.length === 2) return Math.hypot(x - p[0]!, y - p[1]!) <= reach
    for (let i = 0; i + 3 < p.length; i += 2) {
      if (segmentDistance(x, y, p[i]!, p[i + 1]!, p[i + 2]!, p[i + 3]!) <= reach) return true
    }
    return false
  })
}
