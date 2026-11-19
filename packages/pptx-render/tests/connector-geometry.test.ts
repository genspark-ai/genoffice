/**
 * Connector geometry tests:
 * - bent elbow positions (controlled by adj)
 * - curved bezier control point counts
 * - ArrowHead size conversion
 */
import { describe, it, expect } from 'vitest'
import { connectorPoints, connectorBezier } from '../src/preset-geometry'
import { openPptx } from '@genoffice/pptx-engine'
import { buildRenderSlide } from '../src/index'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'

const here = dirname(fileURLToPath(import.meta.url))
const enginePptx = (name: string) =>
  readFileSync(join(here, '..', '..', 'pptx-engine', 'tests', 'fixtures', name))

// ── bent connector elbows ─────────────────────────────────────────────────

describe('connectorPoints: bentConnector3 adj controls elbow position', () => {
  it('adj1=50000 → elbow at horizontal center (default)', () => {
    const pts = connectorPoints('bentConnector3', 100, 60, false, false, { adj1: 50000 })
    // Expected: [0,0, 50,0, 50,60, 100,60]
    expect(pts).toHaveLength(8)
    expect(pts[2]).toBeCloseTo(50, 3) // second elbow x
    expect(pts[4]).toBeCloseTo(50, 3) // third elbow x
  })

  it('adj1=75000 → elbow at 75% of width', () => {
    const pts = connectorPoints('bentConnector3', 100, 60, false, false, { adj1: 75000 })
    expect(pts[2]).toBeCloseTo(75, 3)
    expect(pts[4]).toBeCloseTo(75, 3)
  })

  it('flipV mirrors y coordinates', () => {
    const pts = connectorPoints('bentConnector3', 100, 60, false, true, { adj1: 50000 })
    // flipV: y = h - y, so [0,60, 50,60, 50,0, 100,0]
    expect(pts[1]).toBeCloseTo(60, 3) // start y
    expect(pts[7]).toBeCloseTo(0, 3) // end y
  })

  it('flipH mirrors x coordinates', () => {
    const pts = connectorPoints('bentConnector3', 100, 60, true, false, { adj1: 50000 })
    expect(pts[0]).toBeCloseTo(100, 3) // start x = w - 0
    expect(pts[6]).toBeCloseTo(0, 3) // end x = w - w
  })
})

describe('connectorPoints: bentConnector2 (single elbow)', () => {
  it('3 bend points', () => {
    const pts = connectorPoints('bentConnector2', 100, 60, false, false)
    expect(pts).toHaveLength(6) // 3 * 2
    // [0,0, 100,0, 100,60]
    expect(pts[2]).toBeCloseTo(100, 3)
    expect(pts[3]).toBeCloseTo(0, 3)
    expect(pts[4]).toBeCloseTo(100, 3)
    expect(pts[5]).toBeCloseTo(60, 3)
  })
})

describe('connectorPoints: line/straightConnector', () => {
  it('two endpoints', () => {
    const pts = connectorPoints('line', 100, 60, false, false)
    expect(pts).toHaveLength(4)
    expect(pts).toEqual([0, 0, 100, 60])
  })

  it('correct diagonal after flip', () => {
    const pts = connectorPoints('straightConnector1', 100, 60, true, true)
    expect(pts).toEqual([100, 60, 0, 0])
  })
})

// ── curved connector bezier ───────────────────────────────────────────────

describe('connectorBezier', () => {
  it('2 bend points return empty (no curve)', () => {
    const pts = [0, 0, 100, 60]
    expect(connectorBezier(pts)).toEqual([])
  })

  it('3 bend points (bentConnector2) return 2 bezier segments (12 numbers)', () => {
    const pts = [0, 0, 100, 0, 100, 60]
    const bz = connectorBezier(pts)
    expect(bz).toHaveLength(12) // 2 segments × 6
  })

  it('4 bend points (bentConnector3) return 3 bezier segments (18 numbers)', () => {
    const pts = [0, 0, 50, 0, 50, 60, 100, 60]
    const bz = connectorBezier(pts)
    expect(bz).toHaveLength(18) // 3 segments × 6
  })

  it('last 2 numbers of each segment equal the polyline bend point', () => {
    const pts = [0, 0, 50, 0, 50, 60, 100, 60]
    const bz = connectorBezier(pts)
    // segment 0 endpoint = pts[1] = (50,0)
    expect(bz[4]).toBeCloseTo(50, 5)
    expect(bz[5]).toBeCloseTo(0, 5)
    // segment 1 endpoint = pts[2] = (50,60)
    expect(bz[10]).toBeCloseTo(50, 5)
    expect(bz[11]).toBeCloseTo(60, 5)
  })
})

// ── build-slide ArrowEnd → ArrowEndRender ─────────────────────────────────

describe('buildRenderSlide connector → line.tailEnd/headEnd', () => {
  it('tailEnd triangle → line.tailEnd.type=triangle, sized by med line width', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const el: any = {
      id: 'c1',
      type: 'shape',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      transform: {
        offset: { x: 0, y: 0, cx: 914400, cy: 457200 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      presetGeometry: 'straightConnector1',
      fill: { type: 'none' },
      stroke: {
        fill: { type: 'solid', color: '#000000' },
        width: 19050,
        tailEnd: { type: 'triangle', w: 'med', len: 'med' },
      },
    }
    const rs = buildRenderSlide({ ...slide, elements: [el], decorations: [] }, deck.size, {
      fitWidthPx: 1280,
    })
    const n = rs.nodes[0] as any
    expect(n.line).toBeTruthy()
    expect(n.line.tailEnd).toBeTruthy()
    expect(n.line.tailEnd.type).toBe('triangle')
    // med x 3 = ~3*lw; lw ~= 19050/9525 * scale(1280/...)
    expect(n.line.tailEnd.widthPx).toBeGreaterThan(0)
    expect(n.line.tailEnd.lengthPx).toBeGreaterThan(0)
  })

  it('both headEnd and tailEnd present', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const el: any = {
      id: 'c2',
      type: 'shape',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      transform: {
        offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      presetGeometry: 'bentConnector3',
      adjust: { adj1: 50000 },
      fill: { type: 'none' },
      stroke: {
        fill: { type: 'solid', color: '#FF0000' },
        width: 12700,
        headEnd: { type: 'oval' },
        tailEnd: { type: 'diamond', w: 'lg' },
      },
    }
    const rs = buildRenderSlide({ ...slide, elements: [el], decorations: [] }, deck.size, {
      fitWidthPx: 1280,
    })
    const n = rs.nodes[0] as any
    expect(n.line.headEnd?.type).toBe('oval')
    expect(n.line.tailEnd?.type).toBe('diamond')
    // lg multiplier 5 > med 3
    expect(n.line.tailEnd.widthPx).toBeGreaterThan(n.line.headEnd.widthPx)
  })

  it('curved connector produces bezier control points', async () => {
    const { deck } = await openPptx(enginePptx('01_standard_business.pptx'))
    const slide = deck.slides[0]!
    const el: any = {
      id: 'c3',
      type: 'shape',
      anchor: { spIndex: -1, originalXml: '', range: [0, 0] },
      transform: {
        offset: { x: 0, y: 0, cx: 914400, cy: 457200 },
        rot: 0,
        flipH: false,
        flipV: false,
      },
      presetGeometry: 'curvedConnector3',
      adjust: { adj1: 50000 },
      fill: { type: 'none' },
      stroke: { fill: { type: 'solid', color: '#000000' }, width: 12700 },
    }
    const rs = buildRenderSlide({ ...slide, elements: [el], decorations: [] }, deck.size, {
      fitWidthPx: 1280,
    })
    const n = rs.nodes[0] as any
    expect(n.line).toBeTruthy()
    expect(n.line.bezier).toBeTruthy()
    expect(n.line.bezier.length).toBeGreaterThan(0)
  })
})
