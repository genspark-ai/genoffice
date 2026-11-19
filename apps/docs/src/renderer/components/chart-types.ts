import type { ChartexLayout, NewChart } from '@genoffice/docx-engine'
import type { StringKey } from '../i18n/locale'

/** the chart types Insert > Chart can start, in Word's menu order */
export const CHART_TYPES = [
  'bar',
  'barH',
  'line',
  'area',
  'pie',
  'treemap',
  'sunburst',
  'histogram',
  'pareto',
  'boxWhisker',
  'scatter',
  'waterfall',
  'funnel',
  'stock',
  'surface',
  'radar',
] as const
export type ChartType = (typeof CHART_TYPES)[number]

export const CHART_TYPE_LABELS: Record<ChartType, StringKey> = {
  bar: 'ribbonChartBar',
  barH: 'ribbonChartBarH',
  line: 'ribbonChartLine',
  area: 'ribbonChartArea',
  pie: 'ribbonChartPie',
  scatter: 'ribbonChartScatter',
  stock: 'ribbonChartStock',
  surface: 'ribbonChartSurface',
  radar: 'ribbonChartRadar',
  treemap: 'ribbonChartTreemap',
  sunburst: 'ribbonChartSunburst',
  histogram: 'ribbonChartHistogram',
  pareto: 'ribbonChartPareto',
  boxWhisker: 'ribbonChartBoxWhisker',
  waterfall: 'ribbonChartWaterfall',
  funnel: 'ribbonChartFunnel',
}

/** the types of the 2014 chart extension, with the classic kind each one reads back over */
const EXTENSION_KINDS: Partial<Record<ChartType, NewChart['kind']>> = {
  treemap: 'pie',
  sunburst: 'pie',
  histogram: 'bar',
  pareto: 'bar',
  boxWhisker: 'bar',
  waterfall: 'bar',
  funnel: 'bar',
}

/**
 * Word's Bar is a bar chart lying down; a type of the chart extension is its
 * layout over the nearest classic kind; every other type is a kind of its own.
 */
export function chartKindOf(type: ChartType): Pick<NewChart, 'kind' | 'horizontal' | 'layout'> {
  if (type === 'barH') return { kind: 'bar', horizontal: true }
  const over = EXTENSION_KINDS[type]
  return over ? { kind: over, layout: type as ChartexLayout } : { kind: type as NewChart['kind'] }
}

/** these read one series: a second one in the grid would not be drawn */
export const SINGLE_SERIES_TYPES: ReadonlySet<ChartType> = new Set([
  'treemap',
  'sunburst',
  'histogram',
  'pareto',
  'waterfall',
  'funnel',
])

/** a stock chart reads its series as high, low and close: it needs all three; a fourth is the open */
export const STOCK_MIN_SERIES = 3
export const STOCK_MAX_SERIES = 4
