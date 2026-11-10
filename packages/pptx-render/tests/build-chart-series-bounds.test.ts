import { describe, expect, it } from 'vitest'
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

const N = 130_000
const TIMEOUT = 60_000

const onePointEach = (): ChartModel['series'] =>
  Array.from({ length: N }, (_, i) => ({ values: [i % 97, (i % 31) + 1] }))

const emptyButNamed = (): ChartModel['series'] => {
  const series: ChartModel['series'] = Array.from({ length: N }, () => ({ values: [] }))
  series[0] = { name: 'only', values: [5] }
  return series
}

describe('buildChartNode with an oversized series array', () => {
  it(
    'lays out a line chart without spreading the series array',
    () => {
      const model: ChartModel = {
        kind: 'line',
        categories: [''],
        series: onePointEach(),
        catAxis: { tickLblHidden: true },
      }
      expect(() =>
        buildChartNode('many-lines', 'many-lines', model, box, vp, metrics),
      ).not.toThrow()
    },
    TIMEOUT,
  )

  it(
    'measures a side legend and category count without spreading the series array',
    () => {
      const model: ChartModel = {
        kind: 'bar',
        barDir: 'col',
        categories: [''],
        series: emptyButNamed(),
        legendPos: 'r',
        bar3D: {
          rotX: 15,
          rotY: 20,
          depthPct: 100,
          rAngAx: true,
          gapDepthPct: 150,
          serAxLabels: false,
        },
        catAxis: { tickLblHidden: true },
      }
      expect(() => buildChartNode('many-bars', 'many-bars', model, box, vp, metrics)).not.toThrow()
    },
    TIMEOUT,
  )
})
