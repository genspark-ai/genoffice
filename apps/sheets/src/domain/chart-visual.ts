import { numfmt } from '@univerjs/core'

import { columnLabel, parseAddress, parseRange, rangeCellCount } from './cell-address'

/// Chart visual built from a data range. Structurally a subset of the shared
/// WorkbookVisualObject so both the demo snapshot and the lazy edit journal
/// can hold it (arrays stay mutable for that assignability).
export interface SheetVisualAnchor {
  fromRow: number
  fromColumn: number
  fromRowOffset: number
  fromColumnOffset: number
  toRow: number
  toColumn: number
  toRowOffset: number
  toColumnOffset: number
}

export type ChartSeriesState = ChartSeriesVisualState

export interface SheetVisual {
  id: string
  sheetId: string
  kind: 'chart'
  anchor: SheetVisualAnchor
  chart: ChartVisualState
}

export type ChartGridValue = string | number | boolean | null | undefined

/// Full chart shape shared by the demo snapshot, the lazy journal, and the
/// renderer's WorkbookVisualObject.chart (structurally identical).
export interface ChartSeriesVisualState {
  name: string
  categories: string[]
  values: number[]
  numberFormat?: string | undefined
  /// numCache formatCode of the category (or scatter X) data.
  categoryFormat?: string | undefined
  color?: string | undefined
  trendline?: string | undefined
  valuesRef?: string | undefined
  categoriesRef?: string | undefined
  pointColors?: { index: number; color: string }[] | undefined
  /// Pie: `c:ser/c:explosion`, % of radius every slice moves outward.
  explosionPct?: number | undefined
  /// Pie: per-slice `c:dPt/c:explosion` overrides.
  pointExplosions?: { index: number; pct: number }[] | undefined
}

export interface ChartVisualState {
  chartTypes: string[]
  barDirection?: string | undefined
  title: string
  series: ChartSeriesVisualState[]
  legend?: 'none' | 'right' | 'bottom' | 'top' | 'left' | undefined
  axisTitles?:
    { category?: string | null | undefined; value?: string | null | undefined } | undefined
  dataLabels?: 'none' | 'value' | 'percent' | 'category-percent' | undefined
  dataLabelPosition?: 'center' | 'inside-end' | 'outside-end' | undefined
  dataLabelFormat?: string | undefined
  grouping?: 'clustered' | 'stacked' | 'percentStacked' | 'standard' | undefined
  /// Value-axis major gridlines; absent on pie/doughnut.
  gridlines?: boolean | undefined
  /// Explicit value-axis bounds (`c:scaling`); absent keys mean auto.
  valueAxis?: { min?: number | undefined; max?: number | undefined } | undefined
  /// `c:numFmt` on the category/date axis.
  categoryAxisFormat?: string | undefined
  /// Bar family `c:gapWidth` (percent of one bar width between categories).
  gapWidthPct?: number | undefined
  /// Doughnut `c:holeSize`.
  holeSizePct?: number | undefined
}

/// Numeric category labels (date serials, percents) render through their
/// source number format instead of the raw number.
export function formatCategoryLabel(raw: string, format: string | undefined): string {
  if (format === undefined || format === 'General' || raw.trim() === '') return raw
  const value = Number(raw)
  if (!Number.isFinite(value)) return raw
  try {
    return numfmt.format(format, value, { throws: false })
  } catch {
    return raw
  }
}

/// The axis-level `c:numFmt` wins over the first series' numCache formatCode
/// (Excel applies sourceLinked="0" axis formats over the cached ones).
export function chartCategoryFormat(
  chart: Pick<ChartVisualState, 'categoryAxisFormat' | 'series'>,
): string | undefined {
  const axis = chart.categoryAxisFormat
  if (axis !== undefined && axis !== 'General') return axis
  const series = chart.series[0]?.categoryFormat
  return series !== undefined && series !== 'General' ? series : undefined
}

export interface ScatterAxis {
  min: number
  max: number
  ticks: number[]
}

/// Excel-like auto scale for a scatter axis: explicit `c:scaling`
/// bounds win; otherwise 0 (below-zero data pushes further down) up to a
/// nice-step ceiling. Ticks are 5 evenly spaced values.
export function scatterAxisBounds(
  values: readonly number[],
  explicit?: { min?: number | undefined; max?: number | undefined },
): ScatterAxis {
  const finite = values.filter((value) => Number.isFinite(value))
  const dataMin = finite.length > 0 ? Math.min(...finite) : 0
  const dataMax = finite.length > 0 ? Math.max(...finite) : 1
  const min = explicit?.min ?? (dataMin >= 0 ? 0 : -niceCeiling(-dataMin))
  let max = explicit?.max ?? (dataMax <= 0 ? 0 : niceCeiling(dataMax))
  if (!(max > min)) max = min + 1
  const ticks = [0, 0.25, 0.5, 0.75, 1].map((fraction) => min + fraction * (max - min))
  return { min, max, ticks }
}

/// Smallest 1/2/2.5/5×10^n step whose multiple covers `value` within 9
/// intervals — Excel-ish: 0.152 → 0.16, 0.449 → 0.45, 7 → 7, 130 → 140.
function niceCeiling(value: number): number {
  let exponent = Math.floor(Math.log10(value)) - 1
  for (let guard = 0; guard < 8; guard += 1) {
    for (const base of [1, 2, 2.5, 5]) {
      const step = base * 10 ** exponent
      const intervals = Math.ceil(value / step - 1e-9)
      if (intervals <= 9) return intervals * step
    }
    exponent += 1
  }
  return value
}

/// Excel inserts charts with value labels off; opened single-series
/// bar/column charts upgrade to labelled by default (product decision —
/// diverges from source-file fidelity until the user removes the labels).
export function withDefaultBarLabels(chart: ChartVisualState): ChartVisualState {
  const upgradable =
    chart.chartTypes.length === 1 &&
    chart.chartTypes[0] === 'barChart' &&
    chart.series.length === 1 &&
    chart.grouping !== 'stacked' &&
    chart.grouping !== 'percentStacked' &&
    (chart.dataLabels === undefined || chart.dataLabels === 'none')
  return upgradable ? { ...chart, dataLabels: 'value' } : chart
}

/// Mirrors WorkbookChartEdit sans chartPath, so charts that only live in
/// memory (demo snapshot, session adds) take the same edits as file charts.
export interface ChartStateEdit {
  title?: string | undefined
  chartType?: 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | undefined
  seriesColors?: Record<string, string> | undefined
  pointColors?: Record<string, Record<string, string>> | undefined
  legend?: ChartVisualState['legend']
  dataLabels?: ChartVisualState['dataLabels']
  dataLabelPosition?: ChartVisualState['dataLabelPosition']
  dataLabelFormat?: ChartVisualState['dataLabelFormat']
  axisTitles?: ChartVisualState['axisTitles']
  /// Bar/line/area stacking; 'clustered' means side-by-side.
  grouping?: 'clustered' | 'stacked' | 'percentStacked' | undefined
  gridlines?: boolean | undefined
  /// null resets that bound to auto.
  valueAxis?: { min?: number | null | undefined; max?: number | null | undefined } | undefined
  gapWidthPct?: number | undefined
  holeSizePct?: number | undefined
  /// Pie whole-ring explosion (series 0).
  explosionPct?: number | undefined
  /// Pie per-slice explosion overrides (series 0), point index → %.
  pointExplosions?: Record<string, number> | undefined
  series?:
    | {
        index: number
        name?: string | undefined
        values?: number[] | undefined
        categories?: string[] | undefined
        valuesRef?: string | undefined
        categoriesRef?: string | undefined
      }[]
    | undefined
  /// Full series replacement (Select Data dialog / switch row-column):
  /// every existing series drops, these become the chart's series.
  seriesSet?:
    | {
        name: string
        values: number[]
        categories?: string[] | undefined
        valuesRef?: string | undefined
        categoriesRef?: string | undefined
        color?: string | undefined
      }[]
    | undefined
}

/// Switch Row/Column: cache-based transpose — old series names become the
/// categories, each old category becomes a series (cell references cannot
/// survive a transpose). Null when there are no categories to pivot on.
export function transposeChartSeries(
  series: readonly ChartSeriesVisualState[],
  seriesLabel: (n: number) => string,
): NonNullable<ChartStateEdit['seriesSet']> | null {
  const seriesNames = series.map((entry, index) => entry.name || seriesLabel(index + 1))
  const categories = series[0]?.categories ?? []
  if (categories.length === 0) return null
  return categories.slice(0, 24).map((category, index) => ({
    name: (category || seriesLabel(index + 1)).slice(0, 255),
    values: series.map((entry) => entry.values[index] ?? 0),
    categories: seriesNames,
  }))
}

/// Chart families whose plot-level data labels the save pipeline can write
/// (mirrors xlsx-chart setDataLabels).
const LABELABLE_PLOTS = new Set(['barChart', 'lineChart', 'areaChart', 'pieChart', 'doughnutChart'])

export function chartSupportsDataLabels(chartTypes: readonly string[]): boolean {
  return chartTypes.some((type) => LABELABLE_PLOTS.has(type))
}

/// Single-plot families whose series set the save pipeline can replace
/// (mirrors xlsx-chart replaceSeries).
const SERIES_REPLACEABLE_PLOTS = new Set([
  'barChart',
  'lineChart',
  'areaChart',
  'pieChart',
  'doughnutChart',
  'radarChart',
  'scatterChart',
])

export function chartSupportsSeriesReplace(chartTypes: readonly string[]): boolean {
  return chartTypes.length === 1 && SERIES_REPLACEABLE_PLOTS.has(chartTypes[0] ?? '')
}

export const CHART_EDIT_TYPES: Record<
  string,
  Pick<ChartVisualState, 'chartTypes' | 'barDirection'>
> = {
  column: { chartTypes: ['barChart'], barDirection: 'col' },
  bar: { chartTypes: ['barChart'], barDirection: 'bar' },
  line: { chartTypes: ['lineChart'] },
  area: { chartTypes: ['areaChart'] },
  pie: { chartTypes: ['pieChart'] },
  doughnut: { chartTypes: ['doughnutChart'] },
}

function mergePointColors(
  existing: { index: number; color: string }[] | undefined,
  patch: Record<string, string> | undefined,
): { index: number; color: string }[] | undefined {
  if (!patch) return existing
  const byIndex = new Map((existing ?? []).map((point) => [point.index, point.color]))
  for (const [key, color] of Object.entries(patch)) byIndex.set(Number(key), color)
  return [...byIndex].map(([index, color]) => ({ index, color })).sort((a, b) => a.index - b.index)
}

function mergePointExplosions(
  existing: { index: number; pct: number }[] | undefined,
  patch: Record<string, number> | undefined,
): { index: number; pct: number }[] | undefined {
  if (!patch) return existing
  const byIndex = new Map((existing ?? []).map((point) => [point.index, point.pct]))
  for (const [key, pct] of Object.entries(patch)) byIndex.set(Number(key), pct)
  return [...byIndex].map(([index, pct]) => ({ index, pct })).sort((a, b) => a.index - b.index)
}

/// min/max: number sets the bound, null resets to auto, absent keeps.
function mergeValueAxis(
  existing: ChartVisualState['valueAxis'],
  patch: ChartStateEdit['valueAxis'],
): ChartVisualState['valueAxis'] {
  if (!patch) return existing
  const merged: { min?: number; max?: number } = {}
  const min = patch.min === undefined ? existing?.min : patch.min
  const max = patch.max === undefined ? existing?.max : patch.max
  if (typeof min === 'number') merged.min = min
  if (typeof max === 'number') merged.max = max
  return Object.keys(merged).length === 0 ? undefined : merged
}

/// Applies an edit over a chart state (immutably). Both the on-screen
/// preview of pending file-chart edits and the baked edits of in-memory
/// charts run through here, so they cannot drift apart.
export function applyChartStateEdit(
  chart: ChartVisualState,
  edit: ChartStateEdit | undefined,
): ChartVisualState {
  if (!edit) return chart
  const typeOverride: Partial<Pick<ChartVisualState, 'chartTypes' | 'barDirection'>> =
    edit.chartType === undefined ? {} : (CHART_EDIT_TYPES[edit.chartType] ?? {})
  // 'clustered' is a bar-only token; line/area spell it 'standard'.
  const effectiveTypes = typeOverride.chartTypes ?? chart.chartTypes
  const grouping =
    edit.grouping === 'clustered' && !effectiveTypes.includes('barChart')
      ? 'standard'
      : edit.grouping
  const fallbackCategories = chart.series[0]?.categories ?? []
  // A full replacement drops every existing series (and their per-point
  // styling — matching the save-side rewrite).
  const baseSeries: ChartSeriesVisualState[] = edit.seriesSet
    ? edit.seriesSet.map((entry) => ({
        name: entry.name,
        categories: entry.categories ?? fallbackCategories.slice(0, entry.values.length),
        values: entry.values,
        ...(entry.valuesRef === undefined ? {} : { valuesRef: entry.valuesRef }),
        ...(entry.categoriesRef === undefined ? {} : { categoriesRef: entry.categoriesRef }),
        ...(entry.color === undefined ? {} : { color: entry.color }),
      }))
    : chart.series
  const valueAxis = mergeValueAxis(chart.valueAxis, edit.valueAxis)
  return {
    ...chart,
    ...(edit.title === undefined ? {} : { title: edit.title }),
    ...(edit.legend === undefined ? {} : { legend: edit.legend }),
    ...(edit.dataLabels === undefined ? {} : { dataLabels: edit.dataLabels }),
    ...(edit.dataLabelPosition === undefined ? {} : { dataLabelPosition: edit.dataLabelPosition }),
    ...(edit.dataLabelFormat === undefined ? {} : { dataLabelFormat: edit.dataLabelFormat }),
    ...(edit.axisTitles === undefined
      ? {}
      : { axisTitles: { ...chart.axisTitles, ...edit.axisTitles } }),
    ...(grouping === undefined ? {} : { grouping }),
    ...(edit.gridlines === undefined ? {} : { gridlines: edit.gridlines }),
    ...(edit.valueAxis === undefined ? {} : { valueAxis }),
    ...(edit.gapWidthPct === undefined ? {} : { gapWidthPct: edit.gapWidthPct }),
    ...(edit.holeSizePct === undefined ? {} : { holeSizePct: edit.holeSizePct }),
    ...typeOverride,
    series: baseSeries.map((series, index) => {
      const color = edit.seriesColors?.[String(index)]
      const pointColors = mergePointColors(series.pointColors, edit.pointColors?.[String(index)])
      const pointExplosions =
        index === 0
          ? mergePointExplosions(series.pointExplosions, edit.pointExplosions)
          : series.pointExplosions
      const data = edit.series?.find((entry) => entry.index === index)
      return {
        ...series,
        ...(color === undefined ? {} : { color }),
        ...(pointColors === undefined ? {} : { pointColors }),
        ...(pointExplosions === undefined ? {} : { pointExplosions }),
        ...(index === 0 && edit.explosionPct !== undefined
          ? { explosionPct: edit.explosionPct }
          : {}),
        ...(data?.name === undefined ? {} : { name: data.name }),
        ...(data?.values === undefined ? {} : { values: data.values }),
        ...(data?.categories === undefined ? {} : { categories: data.categories }),
        ...(data?.valuesRef === undefined ? {} : { valuesRef: data.valuesRef }),
        ...(data?.categoriesRef === undefined ? {} : { categoriesRef: data.categoriesRef }),
      }
    }),
  }
}

export interface ParsedChartData {
  /// Series run along rows (selection at least as wide as tall, Excel's
  /// default). In that orientation hasHeaderRow refers to the first
  /// COLUMN (series names), hasCategoryColumn to the first ROW, and each
  /// series' `column` is its row offset within the selection.
  readonly byRow: boolean
  readonly hasHeaderRow: boolean
  readonly hasCategoryColumn: boolean
  readonly categories: string[]
  /// Offset within the selection along the series axis (for range refs).
  readonly series: { name: string; values: number[]; column: number }[]
}

const MAX_CHART_ROWS = 500
const MAX_CHART_SERIES = 12

/// First row with non-numeric labels → header row; first column mostly
/// non-numeric → category column; every other column becomes a series.
/// Excel's default orientation puts the series along the shorter dimension:
/// a selection at least as wide as it is tall charts each ROW as a series.
/// Transposing up front lets the column-wise logic serve both.
export function chartDataFromValues(
  source: readonly (readonly ChartGridValue[])[],
  options?: {
    /// Scatter: the first data vector is the X axis even when numeric,
    /// so it must become the category vector instead of a Y series.
    numericCategoryColumn?: boolean
  },
): ParsedChartData | null {
  const sourceFirstRow = source[0]
  if (!sourceFirstRow || sourceFirstRow.length === 0) return null
  const byRow = sourceFirstRow.length > 1 && source.length <= sourceFirstRow.length
  const grid = byRow
    ? sourceFirstRow.map((_, column) => source.map((row) => row[column] ?? null))
    : source
  const firstRow = grid[0]
  if (!firstRow || firstRow.length === 0) return null
  const isBlank = (v: unknown): boolean => v === null || v === undefined || v === ''
  const isNumeric = (v: unknown): boolean =>
    typeof v === 'number' ||
    (typeof v === 'string' && v.trim() !== '' && Number.isFinite(Number(v)))
  const toNumber = (v: unknown): number =>
    typeof v === 'number' ? v : isNumeric(v) ? Number(String(v).trim()) : 0
  const width = firstRow.length
  // A header row is signalled by non-numeric labels, or by the cross-tab
  // fingerprint Excel also uses: blank corner cell with filled cells across
  // the rest of the first row (numeric headers — years, months — are data
  // otherwise).
  const hasLabelHeader =
    grid.length > 1 &&
    firstRow.some((v, i) => (width === 1 || i > 0) && !isBlank(v) && !isNumeric(v))
  const hasCrossTabHeader =
    grid.length > 1 &&
    width > 1 &&
    isBlank(firstRow[0]) &&
    firstRow.every((v, i) => i === 0 || !isBlank(v))
  const hasHeaderRow = hasLabelHeader || hasCrossTabHeader
  const body = (hasHeaderRow ? grid.slice(1) : grid).slice(0, MAX_CHART_ROWS)
  if (body.length === 0) return null
  // The blank-corner fingerprint marks the first column as the category axis
  // even when its labels are numeric (years down the side).
  const hasCategoryColumn =
    width > 1 &&
    (hasCrossTabHeader ||
      body.some((row) => !isBlank(row[0]) && !isNumeric(row[0])) ||
      (options?.numericCategoryColumn === true && body.some((row) => isNumeric(row[0]))))
  const categories = body.map((row, i) =>
    hasCategoryColumn ? String(row[0] ?? '') : String(i + 1),
  )
  const series: ParsedChartData['series'] = []
  for (let column = hasCategoryColumn ? 1 : 0; column < width; column++) {
    if (!body.some((row) => isNumeric(row[column]))) continue
    const header = hasHeaderRow ? firstRow[column] : null
    series.push({
      name: isBlank(header) ? `Series ${series.length + 1}` : String(header),
      values: body.map((row) => toNumber(row[column])),
      column,
    })
    if (series.length >= MAX_CHART_SERIES) break
  }
  return series.length > 0 ? { byRow, hasHeaderRow, hasCategoryColumn, categories, series } : null
}

/// `'Sheet Name'!$B$2:$B$13` (chart `c:f` style) → sheet name + plain range.
export function splitSheetRef(ref: string): { sheetName: string; range: string } | null {
  const match = /^'?((?:[^'!]|'')+?)'?!(\$?[A-Z]{1,3}\$?[0-9]+(?::\$?[A-Z]{1,3}\$?[0-9]+)?)$/.exec(
    ref.trim(),
  )
  if (!match?.[1] || !match[2]) return null
  return { sheetName: match[1].replace(/''/g, "'"), range: match[2].replace(/\$/g, '') }
}

/// Rewrites the sheet-name qualifier of a chart `c:f` reference, keeping the
/// raw range text (with its $ anchors) untouched.
export function renameRefSheet(ref: string, oldName: string, newName: string): string {
  const split = splitSheetRef(ref)
  if (!split || split.sheetName !== oldName) return ref
  return `'${newName.replace(/'/g, "''")}'!${ref.slice(ref.indexOf('!') + 1)}`
}

export interface CellBounds {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}

/// Does a chart series reference overlap an edited area on that sheet?
/// Drives the live chart↔data sync.
export function refIntersects(ref: string, sheetName: string, bounds: CellBounds): boolean {
  const split = splitSheetRef(ref)
  if (!split || split.sheetName !== sheetName) return false
  try {
    const area = parseRange(split.range)
    return (
      area.startRow <= bounds.endRow &&
      bounds.startRow <= area.endRow &&
      area.startColumn <= bounds.endColumn &&
      bounds.startColumn <= area.endColumn
    )
  } catch {
    return false
  }
}

/// Absolute A1 ref over one column, rows in 0-based coordinates.
function a1RangeRef(sheetName: string, column: number, fromRow: number, toRow: number): string {
  const col = columnLabel(column)
  const name = sheetName.replace(/'/g, "''")
  return `'${name}'!$${col}$${fromRow + 1}:$${col}$${toRow + 1}`
}

/// Absolute A1 ref over one row, columns in 0-based coordinates.
function a1RowRangeRef(
  sheetName: string,
  row: number,
  fromColumn: number,
  toColumn: number,
): string {
  const name = sheetName.replace(/'/g, "''")
  return `'${name}'!$${columnLabel(fromColumn)}$${row + 1}:$${columnLabel(toColumn)}$${row + 1}`
}

export interface BuildChartVisualInput {
  readonly id: string
  readonly sheetId: string
  readonly sheetName: string
  readonly chartType:
    'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'radar' | 'combo'
  readonly dataRange: string
  readonly title?: string | undefined
  readonly anchorCell?: string | undefined
  readonly values: readonly (readonly ChartGridValue[])[]
}

const CHART_KIND_TO_TYPES: Record<BuildChartVisualInput['chartType'], string[]> = {
  column: ['barChart'],
  bar: ['barChart'],
  line: ['lineChart'],
  area: ['areaChart'],
  pie: ['pieChart'],
  doughnut: ['doughnutChart'],
  scatter: ['scatterChart'],
  radar: ['radarChart'],
  // Bar+line combo: the renderer draws the last series as a line (degrades to
  // pure bars with a single series).
  combo: ['barChart', 'lineChart'],
}

/// Shared AI add_chart assembly (demo adapter + lazy visual pipeline): header
/// and category column are detected from the data, matching the ribbon's
/// Insert Chart. The default anchor sits two columns right of the data.
export function buildChartVisual(input: BuildChartVisualInput): SheetVisual {
  const bounds = parseRange(input.dataRange)
  if (rangeCellCount(bounds) > 2000) {
    throw new Error('add_chart dataRange covers more than 2000 cells.')
  }
  const parsed = chartDataFromValues(input.values)
  if (!parsed) throw new Error('The chart dataRange needs at least one numeric column.')
  const anchorBase = input.anchorCell === undefined ? null : parseAddress(input.anchorCell)
  const anchor: SheetVisualAnchor =
    anchorBase === null
      ? {
          fromRow: bounds.startRow,
          fromColumn: bounds.endColumn + 2,
          fromRowOffset: 0,
          fromColumnOffset: 0,
          toRow: bounds.startRow + 15,
          toColumn: bounds.endColumn + 9,
          toRowOffset: 0,
          toColumnOffset: 0,
        }
      : {
          fromRow: anchorBase.row,
          fromColumn: anchorBase.column,
          fromRowOffset: 0,
          fromColumnOffset: 0,
          toRow: anchorBase.row + 15,
          toColumn: anchorBase.column + 7,
          toRowOffset: 0,
          toColumnOffset: 0,
        }
  const dataStartRow = bounds.startRow + (parsed.hasHeaderRow && !parsed.byRow ? 1 : 0)
  // by-row orientation: series names come from the first column, categories
  // from the first row
  const dataStartColumn = bounds.startColumn + (parsed.hasHeaderRow && parsed.byRow ? 1 : 0)
  return {
    id: input.id,
    sheetId: input.sheetId,
    kind: 'chart',
    anchor,
    chart: {
      chartTypes: [...CHART_KIND_TO_TYPES[input.chartType]],
      ...(input.chartType === 'column' || input.chartType === 'bar' || input.chartType === 'combo'
        ? {
            barDirection: input.chartType === 'bar' ? 'bar' : 'col',
            // Newly created bar/column charts label every bar by default
            // (file charts keep whatever the source workbook configured).
            ...(input.chartType === 'combo' ? {} : { dataLabels: 'value' as const }),
          }
        : {}),
      title:
        input.title ??
        (parsed.series.length === 1 && parsed.series[0] ? parsed.series[0].name : 'Chart Title'),
      series: parsed.series.map((series) => ({
        name: series.name,
        categories: parsed.categories,
        values: series.values,
        ...(parsed.byRow
          ? {
              valuesRef: a1RowRangeRef(
                input.sheetName,
                bounds.startRow + series.column,
                dataStartColumn,
                bounds.endColumn,
              ),
              ...(parsed.hasCategoryColumn
                ? {
                    categoriesRef: a1RowRangeRef(
                      input.sheetName,
                      bounds.startRow,
                      dataStartColumn,
                      bounds.endColumn,
                    ),
                  }
                : {}),
            }
          : {
              valuesRef: a1RangeRef(
                input.sheetName,
                series.column + bounds.startColumn,
                dataStartRow,
                bounds.endRow,
              ),
              ...(parsed.hasCategoryColumn
                ? {
                    categoriesRef: a1RangeRef(
                      input.sheetName,
                      bounds.startColumn,
                      dataStartRow,
                      bounds.endRow,
                    ),
                  }
                : {}),
            }),
      })),
    },
  }
}
