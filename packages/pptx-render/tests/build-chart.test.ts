import { describe, it, expect } from 'vitest'
import type { ChartModel } from '@genoffice/pptx-engine'
import { buildChartNode } from '../src/build-chart'
import { HeuristicMetrics } from '../src/metrics'
import { makeViewport } from '../src/coords'

const vp = makeViewport({ cx: 12192000, cy: 6858000 }, 1280)
const metrics = new HeuristicMetrics()
const box = {
  x: 100,
  y: 100,
  w: 600,
  h: 400,
  centerX: 400,
  centerY: 300,
  rotationDeg: 0,
  flipH: false,
  flipV: false,
}

const lineModel: ChartModel = {
  kind: 'line',
  categories: ['2016', '2017', '2018', '2019'],
  series: [
    { name: 'A', color: '#4CAF50', values: [100, 200, 300, 400], smooth: true, marker: true },
    { name: 'B', color: '#2196F3', values: [50, 150, 250, 350], marker: true },
  ],
  legendPos: 't',
  valAxis: { gridColor: '#E6E6E6', gridDash: true, labelColor: '#666666', labelSizePt: 10 },
}

describe('buildChartNode', () => {
  it('builds line chart: polylines within plot, markers, gridlines, legend', () => {
    const node = buildChartNode('r_1', 'el1', lineModel, box, vp, metrics)!
    expect(node.type).toBe('chart')
    expect(node.polylines).toHaveLength(2)
    expect(node.markers.length).toBe(8)
    expect(node.gridLines.length).toBeGreaterThan(2)
    // Legend: 2 swatches + at least 2 texts (more with tick/category labels)
    expect(node.swatches).toHaveLength(2)
    // All polyline points are inside the chart box
    for (const p of node.polylines) {
      for (let i = 0; i < p.points.length; i += 2) {
        expect(p.points[i]).toBeGreaterThanOrEqual(0)
        expect(p.points[i]).toBeLessThanOrEqual(box.w)
        expect(p.points[i + 1]).toBeGreaterThanOrEqual(0)
        expect(p.points[i + 1]).toBeLessThanOrEqual(box.h)
      }
    }
    // Larger values have smaller y
    const a = node.polylines[0]!.points
    expect(a[1]!).toBeGreaterThan(a[a.length - 1]!)
  })

  it('builds clustered bars with per-series colors and non-overlapping x', () => {
    const model: ChartModel = {
      kind: 'bar',
      barDir: 'col',
      grouping: 'clustered',
      gapWidthPct: 50,
      categories: ['a', 'b'],
      series: [
        { name: 's1', color: '#111111', values: [10, 20] },
        { name: 's2', color: '#222222', values: [15, 5] },
      ],
    }
    const node = buildChartNode('r_2', 'el2', model, box, vp, metrics)!
    expect(node.bars).toHaveLength(4)
    // Bars of two series within a category don't overlap (bars order = series-major: s1[a,b] then s2[a,b])
    const s1 = node.bars.filter((b) => b.color === '#111111')
    const s2 = node.bars.filter((b) => b.color === '#222222')
    expect(s2[0]!.x).toBeGreaterThanOrEqual(s1[0]!.x + s1[0]!.w - 0.01)
    // Larger values give taller bars (smaller y)
    expect(s1[1]!.y).toBeLessThan(s1[0]!.y)
  })

  it('builds bar+line combo: bar series as bars, line series as overlaid polyline', () => {
    const model: ChartModel = {
      kind: 'bar',
      barDir: 'col',
      grouping: 'clustered',
      gapWidthPct: 100,
      legendPos: 'b',
      categories: ['Q1', 'Q2', 'Q3'],
      series: [
        { name: 'Sales', color: '#4472C4', values: [10, 20, 30], plotKind: 'bar' },
        { name: 'Growth Rate', color: '#ED7D31', values: [5, 8, 12], plotKind: 'line' },
      ],
    }
    const node = buildChartNode('r_c', 'elc', model, box, vp, metrics)!
    // The bar series yields only 3 bars (not 6), the line series yields 1 polyline
    expect(node.bars).toHaveLength(3)
    expect(node.bars.every((b) => b.color === '#4472C4')).toBe(true)
    expect(node.polylines).toHaveLength(1)
    expect(node.polylines[0]!.color).toBe('#ED7D31')
    expect(node.polylines[0]!.points).toHaveLength(6)
    // The legend includes both series
    expect(node.swatches).toHaveLength(2)
  })

  it('combo dual axis: line maps to the right secondary value-axis range, right axis line and tick labels drawn', () => {
    const model: ChartModel = {
      kind: 'bar',
      barDir: 'col',
      grouping: 'clustered',
      gapWidthPct: 100,
      categories: ['Q1', 'Q2', 'Q3'],
      series: [
        { name: 'Revenue', color: '#4472C4', values: [120, 180, 240], plotKind: 'bar' },
        {
          name: 'Growth',
          color: '#ED7D31',
          values: [10, 20, 30],
          plotKind: 'line',
          secondaryAxis: true,
        },
      ],
      valAxis: { min: 0, max: 250 },
      valAxis2: { min: 0, max: 30 },
    }
    const node = buildChartNode('r_c2', 'elc2', model, box, vp, metrics)!
    // Three axis lines: bottom x axis + left primary value axis + right secondary value axis (vertical, x = plot-area right edge)
    expect(node.axisLines).toHaveLength(3)
    const rightAxis = node.axisLines[2]!
    expect(rightAxis.x1).toBe(rightAxis.x2)
    expect(rightAxis.x1).toBeGreaterThan(node.axisLines[1]!.x1)
    // Secondary tick labels are drawn outside the right axis, including its top value 30 (range 0-30 independent of the left axis 0-250)
    const secLabels = node.labels.filter((l) => l.x > rightAxis.x1)
    expect(secLabels.length).toBeGreaterThanOrEqual(2)
    expect(secLabels.map((l) => l.text)).toContain('30')
    expect(secLabels.map((l) => l.text)).not.toContain('250')
    // The line's last value 30 = secondary full scale → y at the plot-area top (misusing the primary 0-250 would land ~88% down)
    const line = node.polylines[0]!
    const lastY = line.points[line.points.length - 1]!
    expect(lastY).toBeCloseTo(rightAxis.y1, 1)
    // Bars still use the primary range: the tallest bar 240/250 doesn't reach the top
    expect(Math.min(...node.bars.map((b) => b.y))).toBeGreaterThan(rightAxis.y1)
  })

  it('returns null for unsupported kinds (falls back to chip upstream)', () => {
    const unknown: ChartModel = { kind: 'unknown', categories: ['x'], series: [{ values: [1] }] }
    expect(buildChartNode('r', 's', unknown, box, vp, metrics)).toBeNull()
  })

  it('builds stacked columns: one bar per category per series, segments cumulative', () => {
    const model: ChartModel = {
      kind: 'bar',
      barDir: 'col',
      grouping: 'stacked',
      gapWidthPct: 100,
      categories: ['a', 'b'],
      series: [
        { name: 's1', color: '#111111', values: [10, 20] },
        { name: 's2', color: '#222222', values: [30, 40] },
      ],
    }
    const node = buildChartNode('r_3', 'el3', model, box, vp, metrics)!
    expect(node.bars).toHaveLength(4)
    // Both segments of a category share the same x (same bar)
    const catA = node.bars.filter((_b, i) => i % 2 === 0) // the two segments of category a
    const xs = new Set(node.bars.map((b) => Math.round(b.x)))
    expect(xs.size).toBe(2) // two categories → two x values
    // The s2 segment stacks above s1 (smaller y), and s1 bottom = s2 top + s1 height
    const aSegs = node.bars.filter((b) => Math.round(b.x) === [...xs][0])
    expect(aSegs).toHaveLength(2)
    const [seg1, seg2] = aSegs.sort((p, q) => q.y - p.y) // bottom segment first
    expect(seg2!.y + seg2!.h).toBeCloseTo(seg1!.y, 1)
    void catA
  })

  it('percentStacked normalizes each category to 100%', () => {
    const model: ChartModel = {
      kind: 'bar',
      barDir: 'col',
      grouping: 'percentStacked',
      categories: ['a'],
      series: [
        { name: 's1', color: '#111111', values: [25] },
        { name: 's2', color: '#222222', values: [75] },
      ],
    }
    const node = buildChartNode('r_4', 'el4', model, box, vp, metrics)!
    expect(node.bars).toHaveLength(2)
    // The two segments' total height = the whole plot-area height (0..100%)
    const total = node.bars.reduce((a, b) => a + b.h, 0)
    const top = Math.min(...node.bars.map((b) => b.y))
    const bottom = Math.max(...node.bars.map((b) => b.y + b.h))
    expect(total).toBeCloseTo(bottom - top, 1)
    // 25/75 split
    const [h1, h2] = node.bars.map((b) => b.h).sort((a, b) => a - b)
    expect(h2! / h1!).toBeCloseTo(3, 1)
  })

  it('builds pie wedges summing to 360° with per-point colors', () => {
    const pie: ChartModel = {
      kind: 'pie',
      categories: ['x', 'y', 'z'],
      series: [{ values: [1, 1, 2], pointColors: ['#AA0000', undefined, '#0000AA'] }],
      legendPos: 'r',
    }
    const node = buildChartNode('r_5', 'el5', pie, box, vp, metrics)!
    expect(node.wedges).toHaveLength(3)
    const sweep = node.wedges!.reduce((a, w) => a + w.sweepDeg, 0)
    expect(sweep).toBeCloseTo(360, 5)
    expect(node.wedges![2]!.sweepDeg).toBeCloseTo(180, 5)
    expect(node.wedges![0]!.color).toBe('#AA0000')
    expect(node.wedges![2]!.color).toBe('#0000AA')
    // Starts at 12 o'clock (-90°)
    expect(node.wedges![0]!.startDeg).toBeCloseTo(-90, 5)
    // Pie has no center hole
    expect(node.wedges![0]!.innerR).toBe(0)
    // Legend: 3 category swatches
    expect(node.swatches).toHaveLength(3)
  })

  it('pseudo-3D pie with a single 360° slice still emits a filled face path', () => {
    const pie: ChartModel = {
      kind: 'pie',
      pseudo3D: true,
      categories: ['only', 'zero'],
      series: [{ values: [5, 0] }],
    }
    const node = buildChartNode('r_5b', 'el5b', pie, box, vp, metrics)!
    expect(node.wedges ?? []).toHaveLength(0)
    const faces = node.paths!.filter((p) => p.stroke === '#ffffff')
    expect(faces).toHaveLength(1)
    // The full-circle face must be two half arcs, not one degenerate arc
    expect(faces[0]!.d.match(/A /g)).toHaveLength(2)
    expect(faces[0]!.d.startsWith('M ')).toBe(true)
    // Endpoints of the two arcs differ (top of the ellipse vs its opposite point)
    const nums = faces[0]!.d.match(/-?\d+(\.\d+)?/g)!.map(Number)
    // d = M x0 y0 A rx ry 0 1 1 x1 y1 A rx ry 0 1 1 x0 y0 Z → y0 at index 1, y1 at index 8
    expect(nums[1]).not.toBeCloseTo(nums[8]!, 3)
  })

  it('exploded 2D pie: wedge centers offset along mid-angles, radius shrunk to keep the footprint', () => {
    const base: ChartModel = {
      kind: 'pie',
      categories: ['x', 'y'],
      series: [{ values: [1, 1] }],
    }
    const plain = buildChartNode('r_e0', 'ele0', base, box, vp, metrics)!
    const exploded = buildChartNode(
      'r_e1',
      'ele1',
      { ...base, series: [{ values: [1, 1], explosionPct: 25 }] },
      box,
      vp,
      metrics,
    )!
    const r0 = plain.wedges![0]!.outerR
    const [w1, w2] = exploded.wedges!
    // offset = 25% of diameter, radius / (1 + 2·25%)
    expect(w1!.outerR).toBeCloseTo(r0 / 1.5, 5)
    const off = 0.5 * w1!.outerR
    // slice 1 sweeps [-90°, 90°] → mid-angle 0° (right); slice 2 mirrors left
    expect(w1!.cx - plain.wedges![0]!.cx).toBeCloseTo(off, 5)
    expect(w1!.cy).toBeCloseTo(plain.wedges![0]!.cy, 5)
    expect(w2!.cx - plain.wedges![1]!.cx).toBeCloseTo(-off, 5)
  })

  it('exploded pseudo-3D pie: per-point explosion separates only that slice', () => {
    const pie: ChartModel = {
      kind: 'pie',
      pseudo3D: true,
      rotXDeg: 30,
      categories: ['x', 'y'],
      series: [{ values: [1, 1], pointExplosionPct: [25, undefined] }],
    }
    const node = buildChartNode('r_e2', 'ele2', pie, box, vp, metrics)!
    const faces = node.paths!.filter((p) => p.stroke === '#ffffff')
    expect(faces).toHaveLength(2)
    // face path apex = wedge center: "M cx cy L …"
    const apex = (d: string) =>
      d
        .match(/-?\d+(\.\d+)?/g)!
        .slice(0, 2)
        .map(Number)
    const [a1, a2] = [apex(faces[0]!.d), apex(faces[1]!.d)]
    expect(a1[0]).toBeGreaterThan(a2[0]!)
    expect(a1[1]).toBeCloseTo(a2[1]!, 5)
  })

  it('doughnut hole uses holePct of outer radius', () => {
    const dough: ChartModel = {
      kind: 'pie',
      holePct: 50,
      categories: ['x', 'y'],
      series: [{ values: [1, 1] }],
    }
    const node = buildChartNode('r_6', 'el6', dough, box, vp, metrics)!
    const w = node.wedges![0]!
    expect(w.innerR).toBeCloseTo(w.outerR / 2, 5)
  })

  it('builds horizontal bar: bars extend along x axis, first category at bottom, larger values give longer bars', () => {
    const model: ChartModel = {
      kind: 'bar',
      barDir: 'bar',
      grouping: 'clustered',
      categories: ['Alpha', 'Beta'],
      series: [{ name: 's1', color: '#111111', values: [10, 30] }],
      legendPos: 'b',
    }
    const node = buildChartNode('r_h', 'elh', model, box, vp, metrics)!
    expect(node.bars).toHaveLength(2)
    const [b1, b2] = node.bars
    // PowerPoint default (minMax): the first category is at the bottom
    expect(b1!.y).toBeGreaterThan(b2!.y)
    // The value-30 bar is longer than the value-10 bar (in the width direction)
    expect(b2!.w).toBeGreaterThan(b1!.w)
    // Bar height > 0 and within the box
    for (const b of node.bars) {
      expect(b.h).toBeGreaterThan(0)
      expect(b.x + b.w).toBeLessThanOrEqual(box.w)
    }
  })

  it('horizontal bar reversed (orientation maxMin): first category on top', () => {
    const model: ChartModel = {
      kind: 'bar',
      barDir: 'bar',
      categories: ['Alpha', 'Beta'],
      series: [{ values: [10, 30] }],
      catAxis: { reversed: true },
    }
    const node = buildChartNode('r_h2', 'elh2', model, box, vp, metrics)!
    const [b1, b2] = node.bars
    expect(b1!.y).toBeLessThan(b2!.y)
  })

  it('builds scatter: positioned by two value axes, lineMarker draws points and lines by default', () => {
    const model: ChartModel = {
      kind: 'scatter',
      scatterStyle: 'lineMarker',
      categories: [],
      series: [{ name: 's', color: '#333333', xValues: [1, 2, 4], values: [10, 20, 15] }],
      valAxis: { gridColor: '#E6E6E6' },
      catAxis: { min: 0, max: 5 },
    }
    const node = buildChartNode('r_s', 'els', model, box, vp, metrics)!
    expect(node.markers).toHaveLength(3)
    expect(node.polylines).toHaveLength(1)
    // x=1 < x=2 < x=4 map to monotonically increasing pixel x
    const [m1, m2, m3] = node.markers
    expect(m1!.x).toBeLessThan(m2!.x)
    expect(m2!.x).toBeLessThan(m3!.x)
    // The y=20 point sits above (smaller pixel y) the y=10 point
    expect(m2!.y).toBeLessThan(m1!.y)
  })

  it('scatter marker-only style (marker) draws no connecting lines', () => {
    const model: ChartModel = {
      kind: 'scatter',
      scatterStyle: 'marker',
      categories: [],
      series: [{ xValues: [1, 2], values: [1, 2] }],
    }
    const node = buildChartNode('r_s2', 'els2', model, box, vp, metrics)!
    expect(node.markers).toHaveLength(2)
    expect(node.polylines).toHaveLength(0)
  })

  it('builds radar: one closed polygon per series, filled style has semi-transparent fill', () => {
    const model: ChartModel = {
      kind: 'radar',
      radarStyle: 'filled',
      categories: ['Listen', 'Speak', 'Read', 'Write'],
      series: [{ name: 'A', color: '#ED7D31', values: [3, 4, 5, 2] }],
      legendPos: 'b',
    }
    const node = buildChartNode('r_r', 'elr', model, box, vp, metrics)!
    // Grid rings + 1 series polygon (all closed polylines)
    const seriesPoly = node.polylines.filter((p) => p.fill)
    expect(seriesPoly).toHaveLength(1)
    expect(seriesPoly[0]!.closed).toBe(true)
    expect(seriesPoly[0]!.points).toHaveLength(8)
    expect(seriesPoly[0]!.fill).toMatch(/^rgba\(237, 125, 49/)
    // Spokes = 4 radial grid lines
    expect(node.gridLines).toHaveLength(4)
    // All category labels are drawn
    const catLabels = node.labels.filter((l) =>
      ['Listen', 'Speak', 'Read', 'Write'].includes(l.text),
    )
    expect(catLabels).toHaveLength(4)
  })

  it('radar with fewer than 3 vertices falls back to null (upstream shows a chip)', () => {
    const model: ChartModel = {
      kind: 'radar',
      categories: ['a', 'b'],
      series: [{ values: [1, 2] }],
    }
    expect(buildChartNode('r_r2', 'elr2', model, box, vp, metrics)).toBeNull()
  })

  it('builds area: solid polygon closed to the baseline', () => {
    const model: ChartModel = {
      kind: 'area',
      categories: ['a', 'b', 'c'],
      series: [{ name: 's', color: '#70AD47', values: [10, 30, 20] }],
    }
    const node = buildChartNode('r_a', 'ela', model, box, vp, metrics)!
    expect(node.polylines).toHaveLength(1)
    const p = node.polylines[0]!
    expect(p.closed).toBe(true)
    expect(p.fill).toBe('#70AD47')
    // First and last points land on the same baseline y
    expect(p.points[1]).toBeCloseTo(p.points[p.points.length - 1]!, 5)
  })
})

describe('chartSpace default text size reaches every label group', () => {
  const ptToPxAt = (pt: number) => (pt * 96) / 72 // vp scale is 1280/12192000EMU ≈ 96dpi ⇒ px = pt×96/72×(scale)
  it('category + value labels fall back to defaultTextPt, not 10pt', () => {
    const model: ChartModel = {
      kind: 'bar',
      barDir: 'col',
      grouping: 'clustered',
      categories: ['A', 'B'],
      series: [{ name: 'S', values: [1, 2] }],
      defaultTextPt: 18,
    }
    const node = buildChartNode('r_1', 'el1', model, box, vp, metrics)!
    const sizes = new Set(node.labels.map((l) => Math.round(l.fontSizePx)))
    const scale = vp.scale
    expect(sizes.has(Math.round(ptToPxAt(18) * scale))).toBe(true)
    expect(sizes.has(Math.round(ptToPxAt(10) * scale))).toBe(false)
  })
})

describe('sunburst / funnel (chartEx)', () => {
  it('merges a row ending at an ancestor with deeper rows sharing the path', () => {
    // Row 0 ends at "Stem" (its own value); rows 1-2 are leaves under the same "Stem"
    const model: ChartModel = {
      kind: 'sunburst',
      categories: ['', 'Leaf1', 'Leaf2'],
      series: [{ values: [10, 20, 30] }],
      sunburst: {
        levels: [
          ['', 'Leaf1', 'Leaf2'],
          ['Stem', 'Stem', 'Stem'],
        ],
        sizes: [10, 20, 30],
      },
    }
    const node = buildChartNode('r_1', 'el1', model, box, vp, metrics)!
    const stemWedges = node.wedges!.filter(
      (w) => w.innerR === Math.min(...node.wedges!.map((x) => x.innerR)),
    )
    // One merged "Stem" ring segment spanning the full circle, not two siblings
    expect(stemWedges).toHaveLength(1)
    expect(stemWedges[0]!.sweepDeg).toBeCloseTo(360)
    // Two leaf segments on the outer ring (the ancestor's own 10 leaves a gap)
    expect(node.wedges!.length).toBe(3)
  })

  it('funnel draws centered bars with category labels', () => {
    const model: ChartModel = {
      kind: 'funnel',
      categories: ['A', 'B'],
      series: [{ values: [50, 100] }],
      gapWidthPct: 6,
    }
    const node = buildChartNode('r_1', 'el1', model, box, vp, metrics)!
    expect(node.bars).toHaveLength(2)
    // Widest bar is centered: symmetric margins
    const b = node.bars[1]!
    expect(Math.abs(b.x - (node.box.w - (b.x + b.w)))).toBeLessThan(30)
    expect(node.bars[1]!.w).toBeCloseTo(node.bars[0]!.w * 2, 0)
  })
})

describe('horizontal bar series order and labels (python-pptx barH sheets)', () => {
  const barHModel: ChartModel = {
    kind: 'bar',
    barDir: 'bar',
    categories: ['Foo', 'Bar'],
    series: [
      { name: 'Series 1', color: '#4472C4', values: [1.2, 2.3], dataLabels: true },
      { name: 'Series 2', color: '#ED7D31', values: [4.5, 5.6], dataLabels: true },
    ],
    legendPos: 'r',
    dataLabels: true,
  } as any

  it('series 1 sits nearest the category axis (bottom of each group)', () => {
    const node = buildChartNode('r_1', 'el1', barHModel, box, vp, metrics)!
    // two bars per category; within a group the series-1 bar has the LARGER y
    const s1 = node.bars.filter((b) => b.color === '#4472C4')
    const s2 = node.bars.filter((b) => b.color === '#ED7D31')
    expect(s1[0]!.y).toBeGreaterThan(s2[0]!.y)
  })

  it('legend lists series in reverse (Series 2 first)', () => {
    const node = buildChartNode('r_1', 'el1', barHModel, box, vp, metrics)!
    const legendTexts = node.labels.filter((l) => l.text.startsWith('Series')).map((l) => l.text)
    expect(legendTexts[0]).toBe('Series 2')
  })
})

describe('composite data labels (c:showSerName/showCatName)', () => {
  it('prepends series and category names to the value', () => {
    const m: ChartModel = {
      kind: 'bar',
      barDir: 'bar',
      categories: ['Category 1'],
      series: [{ name: 'Series 2', color: '#ED7D31', values: [2.4], dataLabels: true }],
      dataLabels: true,
      dataLabelSerName: true,
      dataLabelCatName: true,
    } as any
    const node = buildChartNode('r_1', 'el1', m, box, vp, metrics)!
    expect(node.labels.some((l) => l.text === 'Series 2, Category 1, 2.4')).toBe(true)
  })
})

describe('size-aware tick cap', () => {
  const mk = (): ChartModel =>
    ({
      kind: 'bar',
      categories: ['A', 'B'],
      series: [{ name: 'S', color: '#4472C4', values: [12.3, 8] }],
    }) as any

  it('a short chart steps the unit up the 1/2/5 ladder (range 13 -> unit 5)', () => {
    const small = { ...box, h: 110 }
    const node = buildChartNode('r_1', 'el1', mk(), small, vp, metrics)!
    const ticks = node.labels.filter((l) => /^\d+$/.test(l.text)).map((l) => Number(l.text))
    expect(Math.max(...ticks)).toBe(15)
    expect(ticks).toContain(5)
    expect(ticks).not.toContain(2)
  })

  it('a tall chart keeps the dense ratio-rule unit', () => {
    const tall = { ...box, h: 900 }
    const node = buildChartNode('r_1', 'el1', mk(), tall, vp, metrics)!
    const ticks = node.labels.filter((l) => /^\d+$/.test(l.text)).map((l) => Number(l.text))
    expect(ticks).toContain(2)
  })
})

describe('chart wave2: palette idx colors, multi-level cats, barH hidden labels', () => {
  it('automatic series colors key off paletteIdx, not document order', () => {
    const m: ChartModel = {
      kind: 'bar',
      categories: ['A'],
      series: [
        { values: [1], paletteIdx: 2 },
        { values: [2], paletteIdx: 0 },
      ],
      themePalette: ['#111111', '#222222', '#333333'],
    } as any
    const node = buildChartNode('r_1', 'el1', m, box, vp, metrics)!
    const colors = node.bars.map((b) => b.color)
    expect(colors[0]).toBe('#333333')
    expect(colors[1]).toBe('#111111')
  })

  it('renders outer-level group labels centered under their spans', () => {
    const m: ChartModel = {
      kind: 'bar',
      categories: ['SF', 'LA', 'NY', 'Albany'],
      series: [{ values: [1, 2, 3, 4] }],
      categoryGroups: [
        { label: 'CA', start: 0 },
        { label: 'NY2', start: 2 },
      ],
    } as any
    const node = buildChartNode('r_1', 'el1', m, box, vp, metrics)!
    const ca = node.labels.find((l) => l.text === 'CA')!
    const ny = node.labels.find((l) => l.text === 'NY2')!
    expect(ca).toBeTruthy()
    expect(ny.x).toBeGreaterThan(ca.x)
    // group row sits below the leaf category labels
    const leaf = node.labels.find((l) => l.text === 'SF')!
    expect(ca.y).toBeGreaterThan(leaf.y)
  })

  it('barH with tickLblPos none draws no category labels', () => {
    const m: ChartModel = {
      kind: 'bar',
      barDir: 'bar',
      categories: ['Category 1', 'Category 2'],
      series: [{ values: [1, 2] }],
      catAxis: { tickLblHidden: true },
    } as any
    const node = buildChartNode('r_1', 'el1', m, box, vp, metrics)!
    expect(node.labels.some((l) => l.text.startsWith('Category'))).toBe(false)
  })
})

describe('legend swatch metrics (PowerPoint ~0.5em squares)', () => {
  it('side legend swatches are half-em and reserve a tight width', () => {
    const node = buildChartNode('r_1', 'el1', lineModel, box, vp, metrics)!
    for (const sw of node.swatches) {
      expect(sw.w).toBeLessThan(20)
      expect(Math.abs(sw.w - sw.h)).toBeLessThan(1)
    }
  })
})
