import { describe, it, expect } from 'vitest'
import type { ChartModel } from '@genoffice/pptx-engine'
import { buildChartNode } from '../src/build-chart'
import type { ChartRenderNode } from '../src/render-tree'
import { HeuristicMetrics } from '../src/metrics'
import { makeViewport, ptToPx } from '../src/coords'
import { presetPath } from '../src/preset-geometry'

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
const texts = (node: { labels: Array<{ text: string }> }) => node.labels.map((l) => l.text)

describe('chart rendering: deleted axes, spacer series, legend order, pie of pie', () => {
  it('horizontal bars: a deleted value axis draws no tick numbers', () => {
    const model: ChartModel = {
      kind: 'bar',
      barDir: 'bar',
      categories: ['A', 'B'],
      series: [{ name: 'S', values: [40, 60] }],
      valAxis: { hidden: true },
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    expect(texts(node).filter((t) => /^\d/.test(t))).toEqual([])
    expect(texts(node)).toContain('A')
  })

  it('bubble/scatter: deleted axes draw no tick numbers', () => {
    const model: ChartModel = {
      kind: 'scatter',
      categories: [],
      series: [{ name: 'S', values: [1, 2, 3], xValues: [1, 2, 3], bubbleSizes: [5, 6, 7] }],
      valAxis: { hidden: true },
      catAxis: { hidden: true },
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    expect(texts(node).filter((t) => /^-?\d/.test(t))).toEqual([])
  })

  it('stacked columns: a noFill spacer series lifts the next series without painting', () => {
    const model: ChartModel = {
      kind: 'bar',
      barDir: 'col',
      grouping: 'stacked',
      categories: ['Q1', 'Q2'],
      series: [
        { name: 'base', values: [32, 36], noFill: true },
        { name: 'delta', values: [4, 2], color: '#FF0000' },
      ],
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    expect(node.bars).toHaveLength(2)
    // the visible bars float above the spacer instead of starting at the axis
    const plotBottom = Math.max(...node.axisLines.map((a) => Math.max(a.y1, a.y2)))
    for (const b of node.bars) expect(b.y + b.h).toBeLessThan(plotBottom - 20)
  })

  it('legend follows legendOrder (deleted entries gone, stack order reversed)', () => {
    const model: ChartModel = {
      kind: 'bar',
      barDir: 'col',
      grouping: 'stacked',
      categories: ['Q1'],
      series: [
        { name: 'CCA', values: [1] },
        { name: 'Empty', values: [2] },
        { name: 'Growth', values: [3], plotKind: 'line' },
      ],
      legendPos: 'r',
      legendOrder: [0, 2],
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    expect(node.swatches).toHaveLength(2)
    expect(texts(node)).toContain('CCA')
    expect(texts(node)).toContain('Growth')
    expect(texts(node)).not.toContain('Empty')
  })

  it('pie of pie: main pie with an "other" slice, secondary pie to the right, two connector lines', () => {
    const model: ChartModel = {
      kind: 'pie',
      categories: ['EXPORT', 'NORTH', 'WEST', 'SOUTH'],
      series: [{ name: 'S', values: [0, 33, 27, 3] }],
      ofPie: { splitPos: 3, secondPieSize: 75, gapWidth: 100 },
      legendPos: 'b',
      dataLabels: true,
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    // main: EXPORT is zero → only the "other" wedge; secondary: three wedges
    expect(node.wedges).toHaveLength(4)
    const [main, ...secondary] = node.wedges!
    expect(main!.sweepDeg).toBeCloseTo(360, 3)
    for (const w of secondary) {
      expect(w.cx).toBeGreaterThan(main!.cx + main!.outerR)
      expect(w.outerR).toBeCloseTo(main!.outerR * 0.75, 3)
    }
    expect(node.polylines).toHaveLength(2)
    expect(node.swatches).toHaveLength(4)
    expect(texts(node)).toContain('63')
  })
})

describe('horizontal-bar legend order, pie-of-pie shares, clustered spacer bars', () => {
  it('clustered horizontal bars skip a noFill series too', () => {
    const model: ChartModel = {
      kind: 'bar',
      barDir: 'bar',
      grouping: 'clustered',
      categories: ['A', 'B'],
      series: [
        { name: 'spacer', values: [5, 6], noFill: true },
        { name: 'S', values: [40, 60] },
      ],
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    expect(node.bars).toHaveLength(2)
  })

  it('horizontal-bar legend with a deleted entry keeps the bottom-up order', () => {
    const model: ChartModel = {
      kind: 'bar',
      barDir: 'bar',
      categories: ['A'],
      series: [
        { name: 'S1', values: [1] },
        { name: 'S2', values: [2] },
        { name: 'S3', values: [3] },
      ],
      legendPos: 'r',
      // engine display order [S3, S2, S1] with position 1 (S2) deleted
      legendOrder: [2, 0],
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    const legend = node.labels.filter((l) => /^S\d$/.test(l.text))
    expect(legend.map((l) => l.text)).toEqual(['S3', 'S1'])
    expect(legend[0]!.y).toBeLessThan(legend[1]!.y)
  })

  it('pie-of-pie percent labels are shares of the whole series', () => {
    const model: ChartModel = {
      kind: 'pie',
      categories: ['A', 'B', 'C', 'D'],
      series: [{ name: 'S', values: [50, 25, 15, 10] }],
      ofPie: { splitPos: 3, secondPieSize: 75, gapWidth: 100 },
      dataLabels: true,
      dataLabelsPct: true,
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    const pct = node.labels.map((l) => l.text).filter((t) => t.endsWith('%'))
    expect(pct.sort()).toEqual(['10%', '15%', '25%', '50%', '50%'].sort())
  })
})

describe('curved arrows stay inside their box', () => {
  const coords = (path: string) => {
    const nums = (path.match(/-?\d+(\.\d+)?/g) ?? []).map(Number)
    return { xs: nums.filter((_, i) => i % 2 === 0), ys: nums.filter((_, i) => i % 2 === 1) }
  }

  it('curvedDownArrow stays inside its box with the head apex at the bottom-right', () => {
    const r = presetPath('curvedDownArrow', 600, 130, { adj1: 6396, adj2: 24795, adj3: 18357 })!
    expect(r.path).toMatch(/^M /)
    const { xs, ys } = coords(r.path!)
    expect(Math.min(...xs)).toBeGreaterThanOrEqual(-6)
    expect(Math.max(...xs)).toBeLessThanOrEqual(606)
    expect(Math.min(...ys)).toBeGreaterThanOrEqual(-6)
    expect(Math.max(...ys)).toBeLessThanOrEqual(136)
    // first point = head apex (x6, h)
    expect(xs[0]!).toBeGreaterThan(560)
    expect(ys[0]!).toBeCloseTo(130, 3)
  })

  it('curvedUpArrow mirrors curvedDownArrow vertically', () => {
    const down = coords(presetPath('curvedDownArrow', 300, 200)!.path!)
    const up = coords(presetPath('curvedUpArrow', 300, 200)!.path!)
    expect(down.ys[0]!).toBeCloseTo(200, 3)
    expect(up.ys[0]!).toBeCloseTo(0, 3)
    expect(up.xs[0]!).toBeCloseTo(down.xs[0]!, 3)
  })
})

describe('chart rendering: typeface, tick label position, axis titles, per-point labels', () => {
  const byText = (node: ChartRenderNode, t: string) => node.labels.find((l) => l.text === t)

  it('stamps the chart typeface on every label; unset keeps the default', () => {
    const base: ChartModel = {
      kind: 'bar',
      categories: ['A', 'B'],
      series: [{ name: 'S', values: [40, 60] }],
      title: 'T',
    }
    const plain = buildChartNode('r', 'e', base, box, vp, metrics)!
    expect(plain.labels.every((l) => l.fontFamily === undefined)).toBe(true)
    const arial = buildChartNode('r', 'e', { ...base, fontFamily: 'Arial' }, box, vp, metrics)!
    expect(arial.labels.length).toBeGreaterThan(0)
    expect(arial.labels.every((l) => l.fontFamily === 'Arial')).toBe(true)
  })

  it('column chart: tickLblPos low keeps category labels at the plot bottom when values cross zero', () => {
    const model: ChartModel = {
      kind: 'bar',
      categories: ['up', 'down'],
      series: [{ name: 'S', values: [24, -19] }],
    }
    const nextTo = buildChartNode('r', 'e', model, box, vp, metrics)!
    const low = buildChartNode(
      'r',
      'e',
      { ...model, catAxis: { tickLblPos: 'low' } },
      box,
      vp,
      metrics,
    )!
    const zeroLine = nextTo.axisLines[0]!.y1
    expect(byText(nextTo, 'down')!.y).toBeLessThan(zeroLine + 30)
    expect(byText(low, 'down')!.y).toBeGreaterThan(zeroLine + 60)
  })

  it('horizontal bars: explicit majorUnit sets the tick step and both axis titles are drawn', () => {
    const model: ChartModel = {
      kind: 'bar',
      barDir: 'bar',
      categories: ['A', 'B', 'C', 'D'],
      series: [{ name: 'S', values: [84, 22, 39, 64] }],
      valAxis: { min: 0, max: 100, majorUnit: 25, title: 'Percent of respondents' },
      catAxis: { title: 'Survey response' },
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    expect(['0', '25', '50', '75', '100'].every((t) => byText(node, t))).toBe(true)
    expect(byText(node, '20')).toBeUndefined()
    const valTitle = byText(node, 'Percent of respondents')!
    const catTitle = byText(node, 'Survey response')!
    expect(catTitle.rotationDeg).toBe(-90)
    // the value title sits below the tick row, the rotated category title left of the labels
    expect(valTitle.y).toBeGreaterThan(byText(node, '0')!.y)
    expect(catTitle.x).toBeLessThan(byText(node, 'A')!.x)
    expect(node.bars[0]!.x).toBeGreaterThan(catTitle.x + 8)
  })

  it('column chart: a points-only series labels just the overridden bar', () => {
    const model: ChartModel = {
      kind: 'bar',
      categories: ['A', 'B', 'C'],
      series: [
        {
          name: 'S',
          values: [13, 27, 31],
          dataLabels: true,
          dLblOnlyPoints: true,
          dLblOverrides: [{ idx: 1, val: true, cat: false, ser: false, pct: false }],
        },
      ],
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    expect(byText(node, '27')).toBeDefined()
    expect(byText(node, '13')).toBeUndefined()
    expect(byText(node, '31')).toBeUndefined()
  })

  it('doughnut: per-point dLbl overrides replace the series flags, size and color', () => {
    const model: ChartModel = {
      kind: 'pie',
      categories: ['New features', 'Debugging', 'Other'],
      series: [
        {
          name: 'S',
          values: [36.9, 20, 9.6],
          dLblOverrides: [
            { idx: 0, val: true, cat: false, ser: false, pct: false, sizePt: 10, color: '#F7F2EA' },
            { idx: 1, val: true, cat: false, ser: false, pct: false, sizePt: 10, color: '#F7F2EA' },
            { idx: 2, hidden: true },
          ],
        },
      ],
      holePct: 55,
      dataLabels: true,
      dataLabelsPct: true,
      dataLabelCatName: true,
      dataLabelNoValue: true,
      dataLabelPt: 18,
    }
    const node = buildChartNode('r', 'e', model, box, vp, metrics)!
    const lbl = byText(node, '36.9')!
    expect(lbl.color).toBe('#F7F2EA')
    expect(lbl.fontSizePx).toBeCloseTo(ptToPx(10, vp.scale), 5)
    expect(byText(node, 'New features, 55%')).toBeUndefined()
    expect(node.labels.some((l) => l.text.includes('Other'))).toBe(false)
    expect(byText(node, '20')).toBeDefined()
  })
})
