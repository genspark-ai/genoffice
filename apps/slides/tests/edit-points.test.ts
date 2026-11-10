import { describe, expect, it } from 'vitest'
import { presetPath, presetPolygon } from '@genoffice/pptx-render/preset-geometry'
import { custGeomXml, parseCustGeom } from '@genoffice/pptx-engine'
import type { ShapeRenderNode } from '@genoffice/pptx-render'
import {
  canEditPoints,
  deleteVertex,
  insertVertex,
  moveControl,
  moveVertex,
  nearestOnPath,
  parseSvgPath,
  shapeToEditablePath,
  toSvgPath,
  vertices,
  type PathCmd,
} from '../src/renderer/edit-points'

const box = { x: 10, y: 20, w: 200, h: 100, rotationDeg: 0, flipH: false, flipV: false }
const node = (over: Partial<ShapeRenderNode>): ShapeRenderNode =>
  ({
    id: 'r_s',
    type: 'shape',
    sourceId: 's',
    box,
    fill: { kind: 'none' },
    ...over,
  }) as ShapeRenderNode

const curve: PathCmd[] = [
  { op: 'M', pts: [0, 0] },
  { op: 'C', pts: [10, 0, 20, 0, 50, 50] },
  { op: 'Q', pts: [80, 100, 100, 100] },
  { op: 'L', pts: [0, 100] },
  { op: 'Z', pts: [] },
]

describe('shapeToEditablePath', () => {
  it('reads pathData, polygonPoints, round rects and plain rects; refuses connectors and layered paths', () => {
    expect(shapeToEditablePath(node({ pathData: 'M 0 0 L 200 0 L 100 100 Z' }))).toEqual([
      { op: 'M', pts: [0, 0] },
      { op: 'L', pts: [200, 0] },
      { op: 'L', pts: [100, 100] },
      { op: 'Z', pts: [] },
    ])
    expect(shapeToEditablePath(node({ polygonPoints: [0, 100, 100, 0, 200, 100] }))).toEqual([
      { op: 'M', pts: [0, 100] },
      { op: 'L', pts: [100, 0] },
      { op: 'L', pts: [200, 100] },
      { op: 'Z', pts: [] },
    ])
    expect(shapeToEditablePath(node({ presetGeometry: 'rect' }))).toHaveLength(5)
    expect(shapeToEditablePath(node({ type: 'text' }))).toHaveLength(5)
    const rr = shapeToEditablePath(node({ presetGeometry: 'roundRect', cornerRadiusPx: 20 }))!
    expect(rr.filter((c) => c.op === 'C')).toHaveLength(4)
    expect(rr[0]).toEqual({ op: 'M', pts: [20, 0] })
    expect(shapeToEditablePath(node({ line: { points: [0, 0, 1, 1] } }))).toBeNull()
    expect(shapeToEditablePath(node({ fillPathData: 'M 0 0', strokePathData: 'M 0 0' }))).toBeNull()
    expect(shapeToEditablePath(node({ presetGeometry: 'gear6' }))).toBeNull()
    expect(shapeToEditablePath(node({ pathData: 'M 0 0 A 1 1 0 0 0 1 1' }))).toBeNull()
  })

  it('pills and fully rounded squares drop their zero-length sides', () => {
    const circle = shapeToEditablePath(
      node({
        box: { ...box, w: 100, h: 100 } as ShapeRenderNode['box'],
        presetGeometry: 'roundRect',
        cornerRadiusPx: 50,
      }),
    )!
    expect(circle.filter((c) => c.op === 'L')).toHaveLength(0)
    expect(vertices(circle).map((v) => [v.x, v.y])).toEqual([
      [50, 0],
      [100, 50],
      [50, 100],
      [0, 50],
    ])
    const pill = shapeToEditablePath(node({ presetGeometry: 'roundRect', cornerRadiusPx: 50 }))!
    expect(pill.filter((c) => c.op === 'L')).toHaveLength(2)
    const pts = vertices(pill).map((v) => `${v.x},${v.y}`)
    expect(new Set(pts).size).toBe(pts.length)
    expect(pts).toHaveLength(6)
  })

  it('canEditPoints gates on node type and geometry', () => {
    expect(canEditPoints(node({ presetGeometry: 'ellipse', pathData: 'M 0 0 L 1 1 Z' }))).toBe(true)
    expect(canEditPoints({ type: 'picture' })).toBe(false)
    expect(canEditPoints(undefined)).toBe(false)
    expect(canEditPoints(node({ line: { points: [] } } as Partial<ShapeRenderNode>))).toBe(false)
  })
})

describe('vertices', () => {
  it('lists end points with their in/out controls', () => {
    const v = vertices(curve)
    expect(v.map((p) => [p.x, p.y])).toEqual([
      [0, 0],
      [50, 50],
      [100, 100],
      [0, 100],
    ])
    expect(v[0]!.cIn).toBeUndefined()
    expect(v[0]!.cOut).toMatchObject({ x: 10, y: 0, cmd: 1, i: 0 })
    expect(v[1]!.cIn).toMatchObject({ x: 20, y: 0, cmd: 1, i: 2 })
    expect(v[1]!.cOut).toMatchObject({ x: 80, y: 100, cmd: 2, i: 0 })
    expect(v[2]!.cIn).toMatchObject({ x: 80, y: 100, cmd: 2, i: 0 })
    expect(v[2]!.cOut).toBeUndefined()
  })

  it('folds an explicit closing point onto the subpath start', () => {
    const v = vertices([
      { op: 'M', pts: [0, 0] },
      { op: 'L', pts: [10, 0] },
      { op: 'L', pts: [10, 10] },
      { op: 'C', pts: [5, 10, 0, 5, 0, 0] },
      { op: 'Z', pts: [] },
    ])
    expect(v).toHaveLength(3)
    expect(v[0]!.dup).toBe(3)
    expect(v[0]!.cIn).toMatchObject({ x: 0, y: 5, cmd: 3 })
  })
})

describe('moveVertex / moveControl', () => {
  it('drags the adjacent controls with the vertex and leaves the rest alone', () => {
    const v = vertices(curve)[1]!
    const out = moveVertex(curve, v, 5, -5)
    expect(out[1]!.pts).toEqual([10, 0, 25, -5, 55, 45])
    expect(out[2]!.pts).toEqual([85, 95, 100, 100])
    expect(out[0]).toEqual(curve[0])
    expect(curve[1]!.pts[4]).toBe(50)
  })

  it('moves a folded start together with its closing point', () => {
    const cmds: PathCmd[] = [
      { op: 'M', pts: [0, 0] },
      { op: 'L', pts: [10, 0] },
      { op: 'L', pts: [10, 10] },
      { op: 'L', pts: [0, 0] },
      { op: 'Z', pts: [] },
    ]
    const out = moveVertex(cmds, vertices(cmds)[0]!, 1, 2)
    expect(out[0]!.pts).toEqual([1, 2])
    expect(out[3]!.pts).toEqual([1, 2])
  })

  it('moveControl moves only the addressed control', () => {
    const v = vertices(curve)[1]!
    const out = moveControl(curve, v, 'in', 1, 1)
    expect(out[1]!.pts).toEqual([10, 0, 21, 1, 50, 50])
    expect(moveControl(curve, v, 'out', 1, 1)[2]!.pts).toEqual([81, 101, 100, 100])
    expect(moveControl(curve, vertices(curve)[3]!, 'out', 1, 1)).toBe(curve)
  })
})

describe('deleteVertex / insertVertex', () => {
  it('removes a vertex, keeps at least three, and re-homes a deleted start', () => {
    expect(deleteVertex(curve, vertices(curve)[3]!)).toEqual(curve.slice(0, 3).concat(curve[4]!))
    const moved = deleteVertex(curve, vertices(curve)[0]!)!
    expect(moved[0]).toEqual({ op: 'M', pts: [50, 50] })
    expect(moved).toHaveLength(4)
    const tri = shapeToEditablePath(node({ polygonPoints: [0, 0, 10, 0, 0, 10] }))!
    expect(deleteVertex(tri, vertices(tri)[1]!)).toBeNull()
  })

  it('splits lines, quads, cubics and the implicit closing line', () => {
    const line = insertVertex(curve, 3, 0.5)
    expect(line[3]).toEqual({ op: 'L', pts: [50, 100] })
    expect(line[4]).toEqual({ op: 'L', pts: [0, 100] })
    const cubic = insertVertex(curve, 1, 0.5)
    expect(cubic[1]!.op).toBe('C')
    expect(cubic[2]!.op).toBe('C')
    expect(cubic[2]!.pts.slice(4)).toEqual([50, 50])
    const quad = insertVertex(curve, 2, 0.5)
    expect(quad[2]!.op).toBe('Q')
    expect(quad[3]!.pts.slice(2)).toEqual([100, 100])
    const closing = insertVertex(curve, 4, 0.5)
    expect(closing[4]).toEqual({ op: 'L', pts: [0, 50] })
    expect(closing[5]!.op).toBe('Z')
  })

  it('nearestOnPath finds the closest segment and parameter', () => {
    const hit = nearestOnPath(curve, 50, 101)!
    expect(hit.cmd).toBe(3)
    expect(hit.t).toBeCloseTo(0.5, 1)
    expect(hit.dist).toBeCloseTo(1, 1)
    const onClose = nearestOnPath(curve, -1, 50)!
    expect(onClose.cmd).toBe(4)
  })
})

describe('serialization', () => {
  it('parses and prints absolute M/L/C/Q/Z paths', () => {
    const d = 'M 0 0 C 10 0 20 0 50 50 Q 80 100 100 100 L 0 100 Z'
    expect(toSvgPath(parseSvgPath(d)!)).toBe(d)
    expect(parseSvgPath('L 0 0')).toBeNull()
    expect(parseSvgPath('M 0 x')).toBeNull()
  })

  it('a converted preset survives the custGeom write and re-parse', () => {
    const cx = 1828800
    const cy = 914400
    const scale = cx / box.w
    const check = (cmds: PathCmd[], expected: string) => {
      const emu = cmds.map((c) => ({ op: c.op, pts: c.pts.map((v) => v * scale) }))
      const xml = custGeomXml({ w: cx, h: cy, cmds: emu })
      const geo = parseCustGeom(`<p:spPr>${xml}</p:spPr>`, cx, cy)!
      const back = parseSvgPath(geo.path!)!.map((c) => ({
        op: c.op,
        pts: c.pts.map((v, i) => v * (i % 2 === 0 ? box.w : box.h)),
      }))
      const ref = parseSvgPath(expected)!
      expect(back.map((c) => c.op)).toEqual(ref.map((c) => c.op))
      back.forEach((c, i) => c.pts.forEach((v, k) => expect(v).toBeCloseTo(ref[i]!.pts[k]!, 1)))
    }
    const poly = presetPolygon('hexagon', box.w, box.h)!
    check(
      shapeToEditablePath(node({ polygonPoints: poly }))!,
      toSvgPath(shapeToEditablePath(node({ polygonPoints: poly }))!),
    )
    const heart = presetPath('heart', box.w, box.h)!.path!
    check(shapeToEditablePath(node({ pathData: heart }))!, heart)
  })
})
