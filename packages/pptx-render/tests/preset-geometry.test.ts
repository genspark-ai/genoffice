import { describe, it, expect } from 'vitest'
import {
  isConnectorPreset,
  connectorPoints,
  presetPolygon,
  presetPath,
  isPillPreset,
} from '../src/preset-geometry'

describe('connectorPoints', () => {
  it('straight connector runs corner to corner; flip mirrors direction', () => {
    expect(connectorPoints('straightConnector1', 100, 50, false, false)).toEqual([0, 0, 100, 50])
    expect(connectorPoints('line', 100, 50, true, false)).toEqual([100, 0, 0, 50])
    expect(connectorPoints('line', 100, 50, false, true)).toEqual([0, 50, 100, 0])
  })

  it('bentConnector3 has two elbows at adj1 fraction of width', () => {
    const pts = connectorPoints('bentConnector3', 200, 100, false, false, { adj1: 25000 })
    expect(pts).toEqual([0, 0, 50, 0, 50, 100, 200, 100])
  })

  it('bentConnector2 has one elbow', () => {
    expect(connectorPoints('bentConnector2', 200, 100, false, false)).toEqual([
      0, 0, 200, 0, 200, 100,
    ])
  })

  it('isConnectorPreset recognizes connector presets only', () => {
    for (const p of ['line', 'straightConnector1', 'bentConnector3', 'curvedConnector4']) {
      expect(isConnectorPreset(p)).toBe(true)
    }
    for (const p of ['rect', 'ellipse', 'triangle', undefined]) {
      expect(isConnectorPreset(p)).toBe(false)
    }
  })
})

describe('presetPolygon', () => {
  const inBounds = (pts: number[], w: number, h: number) => {
    for (let i = 0; i < pts.length; i += 2) {
      expect(pts[i]!).toBeGreaterThanOrEqual(-0.01)
      expect(pts[i]!).toBeLessThanOrEqual(w + 0.01)
      expect(pts[i + 1]!).toBeGreaterThanOrEqual(-0.01)
      expect(pts[i + 1]!).toBeLessThanOrEqual(h + 0.01)
    }
  }

  it('triangle apex follows adj', () => {
    expect(presetPolygon('triangle', 100, 80)).toEqual([50, 0, 100, 80, 0, 80])
    expect(presetPolygon('triangle', 100, 80, { adj: 0 })).toEqual([0, 0, 100, 80, 0, 80])
  })

  it('diamond / flowChartDecision are midpoint quads', () => {
    expect(presetPolygon('diamond', 100, 60)).toEqual([50, 0, 100, 30, 50, 60, 0, 30])
    expect(presetPolygon('flowChartDecision', 100, 60)).toEqual(presetPolygon('diamond', 100, 60))
  })

  it('rightArrow head/shaft follow adj1/adj2 and stay in bounds', () => {
    const pts = presetPolygon('rightArrow', 200, 100)! // defaults adj1=adj2=50000
    inBounds(pts, 200, 100)
    // The arrow tip is at the right-edge midpoint
    expect(pts).toContain(200)
    // head length = min(w, ss*0.5) = 50 → head starts at x=150
    expect(pts[2]).toBe(150)
    // shaft thickness = h*0.5 → top edge y=25
    expect(pts[1]).toBe(25)
  })

  it('chevron / homePlate / hexagon / star5 stay within bounds', () => {
    for (const p of [
      'chevron',
      'homePlate',
      'hexagon',
      'star5',
      'parallelogram',
      'trapezoid',
      'octagon',
      'plus',
    ]) {
      const pts = presetPolygon(p, 120, 90)
      expect(pts, p).toBeTruthy()
      inBounds(pts!, 120, 90)
      expect(pts!.length).toBeGreaterThanOrEqual(6)
    }
  })

  it('unsupported presets return null (rect fallback upstream)', () => {
    expect(presetPolygon('rect', 100, 100)).toBeNull()
    expect(presetPolygon('cloudCallout', 100, 100)).toBeNull()
    expect(presetPolygon(undefined, 100, 100)).toBeNull()
  })

  it('pill presets flagged for cornerRadius rendering', () => {
    expect(isPillPreset('flowChartTerminator')).toBe(true)
    expect(isPillPreset('roundRect')).toBe(false)
  })

  it('snip1Rect cuts top-right corner by adj fraction of short side', () => {
    expect(presetPolygon('snip1Rect', 200, 100, { adj: 20000 })).toEqual([
      0, 0, 180, 0, 200, 20, 200, 100, 0, 100,
    ])
  })

  it('notchedRightArrow has tail notch at mid height', () => {
    const pts = presetPolygon('notchedRightArrow', 200, 100)!
    inBounds(pts, 200, 100)
    // The last point is the notch (x>0, y=h/2)
    expect(pts[pts.length - 1]).toBe(50)
    expect(pts[pts.length - 2]).toBeGreaterThan(0)
  })

  it('new polygon presets stay in bounds and are non-trivial', () => {
    for (const p of [
      'snip2SameRect',
      'snip2DiagRect',
      'halfFrame',
      'corner',
      'diagStripe',
      'lightningBolt',
      'flowChartPreparation',
      'flowChartManualInput',
      'flowChartManualOperation',
      'flowChartOffpageConnector',
      'flowChartExtract',
      'flowChartMerge',
      'flowChartCollate',
      'gear6',
      'funnel',
      'quadArrow',
      'bentArrow',
      'irregularSeal1',
      'irregularSeal2',
      'star7',
      'star10',
      'star12',
      'star16',
      'star24',
      'star32',
    ]) {
      const pts = presetPolygon(p, 120, 90)
      expect(pts, p).toBeTruthy()
      inBounds(pts!, 120, 90)
      expect(pts!.length, p).toBeGreaterThanOrEqual(6)
    }
  })

  it('star7 has 14 vertices', () => {
    expect(presetPolygon('star7', 100, 100)!.length).toBe(28)
  })

  it('wedgeRectCallout default puts tail on bottom edge toward tip', () => {
    const pts = presetPolygon('wedgeRectCallout', 100, 100)!
    // tip = (50-20.833, 50+62.5) = (29.17, 112.5): may exceed the box (tail tip)
    expect(pts).toContain(112.5)
    expect(pts.length).toBe(14)
  })
})

describe('presetPath', () => {
  const nums = (d: string) =>
    d
      .split(' ')
      .filter((t) => Number.isFinite(Number(t)))
      .map(Number)
  const inBox = (d: string, w: number, h: number, slack = 0.01) => {
    const v = nums(d)
    for (let i = 0; i < v.length; i += 2) {
      expect(v[i]!).toBeGreaterThanOrEqual(-w * 0.2 - slack)
      expect(v[i]!).toBeLessThanOrEqual(w * 1.2 + slack)
      expect(v[i + 1]!).toBeGreaterThanOrEqual(-h * 0.2 - slack)
      expect(v[i + 1]!).toBeLessThanOrEqual(h * 1.2 + slack)
    }
  }

  it('pie default sweeps 0°→270° from center', () => {
    const r = presetPath('pie', 100, 100)!
    expect(r.path).toMatch(/^M 50 50 L 100 50 C /)
    expect(r.path).toMatch(/Z$/)
  })

  it('arc yields stroke-only arc + fill-only pie', () => {
    const r = presetPath('arc', 100, 100)!
    expect(r.strokePath).toMatch(/^M /)
    expect(r.strokePath).not.toMatch(/Z/)
    expect(r.fillPath).toMatch(/^M 50 50/)
    expect(r.fillPath).toMatch(/Z$/)
  })

  it('donut has two subpaths (outer + reversed hole)', () => {
    const r = presetPath('donut', 100, 100)!
    expect(r.path!.match(/M /g)!.length).toBe(2)
    inBox(r.path!, 100, 100)
  })

  it('brackets/braces are stroke-only', () => {
    for (const p of ['leftBracket', 'rightBracket', 'leftBrace', 'rightBrace']) {
      const r = presetPath(p, 40, 120)!
      expect(r.path, p).toBeUndefined()
      expect(r.strokePath, p).toMatch(/^M /)
      inBox(r.strokePath!, 40, 120)
    }
  })

  it('cube/can/bevel/foldedCorner carry decorative strokePath', () => {
    for (const p of [
      'cube',
      'can',
      'bevel',
      'foldedCorner',
      'smileyFace',
      'flowChartPredefinedProcess',
    ]) {
      const r = presetPath(p, 120, 90)!
      expect(r.path, p).toMatch(/^M /)
      expect(r.strokePath, p).toMatch(/^M /)
      inBox(r.path!, 120, 90)
    }
  })

  it('curved/blob presets stay near box', () => {
    for (const p of [
      'heart',
      'moon',
      'sun',
      'cloud',
      'teardrop',
      'plaque',
      'noSmoking',
      'ribbon',
      'ribbon2',
      'wave',
      'doubleWave',
      'uturnArrow',
      'curvedRightArrow',
      'stripedRightArrow',
      'chord',
      'blockArc',
      'frame',
      'round1Rect',
      'round2SameRect',
      'round2DiagRect',
      'snipRoundRect',
      'flowChartDocument',
      'flowChartMultidocument',
      'flowChartConnector',
      'flowChartOr',
      'flowChartSummingJunction',
      'flowChartSort',
      'flowChartDelay',
      'flowChartDisplay',
      'flowChartPunchedTape',
      'flowChartMagneticDisk',
      'flowChartMagneticDrum',
      'flowChartInternalStorage',
    ]) {
      const r = presetPath(p, 120, 90)
      expect(r, p).toBeTruthy()
      const d = r!.path ?? r!.strokePath!
      expect(d, p).toMatch(/^M /)
      inBox(d, 120, 90)
    }
  })

  it('callout paths include tail toward tip', () => {
    for (const p of ['wedgeRoundRectCallout', 'wedgeEllipseCallout', 'cloudCallout']) {
      const r = presetPath(p, 100, 100)!
      expect(r.path, p).toMatch(/^M /)
      // Default tip is below the box → some point has y>h
      expect(Math.max(...nums(r.path!).filter((_, i) => i % 2 === 1)), p).toBeGreaterThan(100)
    }
  })

  it('blockArc respects inner radius adj3', () => {
    const r = presetPath('blockArc', 100, 100, { adj1: 10800000, adj2: 0, adj3: 50000 })!
    const v = nums(r.path!)
    // inner arc radius = 50-50 = 0 → all points converge within the centerline
    for (let i = 0; i < v.length; i += 2) expect(v[i + 1]!).toBeLessThanOrEqual(50.01)
  })

  it('unknown presets return null', () => {
    expect(presetPath('rect', 100, 100)).toBeNull()
    expect(presetPath('ellipse', 100, 100)).toBeNull()
    expect(presetPath(undefined, 100, 100)).toBeNull()
  })
})
