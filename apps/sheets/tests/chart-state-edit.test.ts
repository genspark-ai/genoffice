import { describe, expect, it } from 'vitest'

import {
  applyChartStateEdit,
  chartCategoryFormat,
  formatCategoryLabel,
  refIntersects,
  scatterAxisBounds,
  splitSheetRef,
  withDefaultBarLabels,
  type ChartVisualState,
  valueAxisScale,
} from '../src/domain/chart-visual'

const base = (): ChartVisualState => ({
  chartTypes: ['barChart'],
  barDirection: 'col',
  title: 'Sales',
  gridlines: true,
  valueAxis: { min: 0, max: 100 },
  series: [
    {
      name: 'S1',
      categories: ['a', 'b'],
      values: [1, 2],
      valuesRef: 'Sheet1!$B$2:$B$3',
      color: '#4472C4',
      pointColors: [{ index: 0, color: '#ED7D31' }],
    },
    { name: 'S2', categories: ['a', 'b'], values: [3, 4] },
  ],
})

describe('applyChartStateEdit', () => {
  it('merges value-axis bounds; null resets one bound to auto', () => {
    const withMax = applyChartStateEdit(base(), { valueAxis: { max: 200 } })
    expect(withMax.valueAxis).toEqual({ min: 0, max: 200 })
    const autoMin = applyChartStateEdit(withMax, { valueAxis: { min: null } })
    expect(autoMin.valueAxis).toEqual({ max: 200 })
    const allAuto = applyChartStateEdit(autoMin, { valueAxis: { max: null } })
    expect(allAuto.valueAxis).toBeUndefined()
  })

  it('applies grouping, spelling clustered as standard off the bar family', () => {
    const stacked = applyChartStateEdit(base(), { grouping: 'stacked' })
    expect(stacked.grouping).toBe('stacked')
    const barClustered = applyChartStateEdit(stacked, { grouping: 'clustered' })
    expect(barClustered.grouping).toBe('clustered')
    const lineClustered = applyChartStateEdit(stacked, { chartType: 'line', grouping: 'clustered' })
    expect(lineClustered.grouping).toBe('standard')
    expect(lineClustered.chartTypes).toEqual(['lineChart'])
  })

  it('replaces every series through seriesSet and drops per-point styling', () => {
    const next = applyChartStateEdit(base(), {
      seriesSet: [{ name: 'N1', values: [9, 8], categories: ['x', 'y'], color: '#70AD47' }],
    })
    expect(next.series).toHaveLength(1)
    expect(next.series[0]).toMatchObject({ name: 'N1', values: [9, 8], color: '#70AD47' })
    expect(next.series[0]?.pointColors).toBeUndefined()
    expect(next.series[0]?.valuesRef).toBeUndefined()
  })

  it('seriesSet without categories falls back to the current primary categories', () => {
    const next = applyChartStateEdit(base(), { seriesSet: [{ name: 'N1', values: [9, 8] }] })
    expect(next.series[0]?.categories).toEqual(['a', 'b'])
  })

  it('drops the parsed outer-level groups when an edit replaces the categories', () => {
    const chart = base()
    const first = chart.series[0]
    if (first) first.categoryGroups = [{ label: '2008', start: 0, end: 2 }]
    const kept = applyChartStateEdit(chart, { series: [{ index: 0, values: [5, 6] }] })
    expect(kept.series[0]?.categoryGroups).toEqual([{ label: '2008', start: 0, end: 2 }])
    const replaced = applyChartStateEdit(chart, {
      series: [{ index: 0, categories: ['x', 'y'] }],
    })
    expect(replaced.series[0]?.categoryGroups).toBeUndefined()
  })

  it('drops stale blank markers when an edit replaces the values', () => {
    const chart = base()
    const first = chart.series[0]
    if (first) first.blanks = [1]
    const kept = applyChartStateEdit(chart, { series: [{ index: 0, name: 'R' }] })
    expect(kept.series[0]?.blanks).toEqual([1])
    const replaced = applyChartStateEdit(chart, { series: [{ index: 0, values: [5, 6] }] })
    expect(replaced.series[0]?.blanks).toBeUndefined()
  })

  it('merges slice explosions per point over the series default', () => {
    const pie = applyChartStateEdit(base(), { chartType: 'pie', explosionPct: 10 })
    expect(pie.chartTypes).toEqual(['pieChart'])
    expect(pie.series[0]?.explosionPct).toBe(10)
    const oneSlice = applyChartStateEdit(pie, { pointExplosions: { '1': 40 } })
    expect(oneSlice.series[0]?.pointExplosions).toEqual([{ index: 1, pct: 40 }])
    const another = applyChartStateEdit(oneSlice, { pointExplosions: { '0': 20 } })
    expect(another.series[0]?.pointExplosions).toEqual([
      { index: 0, pct: 20 },
      { index: 1, pct: 40 },
    ])
  })

  it('keeps gap width, hole size, and gridlines edits', () => {
    const next = applyChartStateEdit(base(), {
      gridlines: false,
      gapWidthPct: 60,
      holeSizePct: 70,
    })
    expect(next.gridlines).toBe(false)
    expect(next.gapWidthPct).toBe(60)
    expect(next.holeSizePct).toBe(70)
  })
})

describe('withDefaultBarLabels', () => {
  const single = (): ChartVisualState => {
    const chart = base()
    return { ...chart, series: chart.series.slice(0, 1) }
  }

  it('upgrades a single-series bar chart only when the file has no dLbls', () => {
    // Explicit showVal="0" must stay off — fidelity beats the product default.
    expect(withDefaultBarLabels({ ...single(), dataLabels: 'none' }).dataLabels).toBe('none')
    expect(withDefaultBarLabels(single()).dataLabels).toBe('value')
  })

  it('keeps an explicit label mode', () => {
    const percent = { ...single(), dataLabels: 'percent' as const }
    expect(withDefaultBarLabels(percent)).toBe(percent)
  })

  it('leaves multi-series, stacked, combo, and non-bar charts alone', () => {
    expect(withDefaultBarLabels({ ...base(), dataLabels: 'none' }).dataLabels).toBe('none')
    const stacked = { ...single(), grouping: 'stacked' as const, dataLabels: 'none' as const }
    expect(withDefaultBarLabels(stacked).dataLabels).toBe('none')
    const combo = {
      ...single(),
      chartTypes: ['barChart', 'lineChart'],
      dataLabels: 'none' as const,
    }
    expect(withDefaultBarLabels(combo).dataLabels).toBe('none')
    const pie = { ...single(), chartTypes: ['pieChart'], dataLabels: 'none' as const }
    expect(withDefaultBarLabels(pie).dataLabels).toBe('none')
  })
})

describe('chart data-sync ref helpers', () => {
  it('splits quoted and bare sheet references', () => {
    expect(splitSheetRef("'My Sheet'!$B$2:$B$13")).toEqual({
      sheetName: 'My Sheet',
      range: 'B2:B13',
    })
    expect(splitSheetRef('Sheet1!$A$1')).toEqual({ sheetName: 'Sheet1', range: 'A1' })
    expect(splitSheetRef("'It''s'!C3:C9")).toEqual({ sheetName: "It's", range: 'C3:C9' })
    expect(splitSheetRef('not-a-ref')).toBeNull()
  })

  it('detects intersection between a series ref and edited bounds', () => {
    const ref = "'Data'!$B$2:$B$13"
    // B2:B13 is rows 1-12, column 1 (0-based)
    expect(
      refIntersects(ref, 'Data', { startRow: 5, endRow: 5, startColumn: 1, endColumn: 1 }),
    ).toBe(true)
    expect(
      refIntersects(ref, 'Data', { startRow: 0, endRow: 0, startColumn: 1, endColumn: 1 }),
    ).toBe(false)
    expect(
      refIntersects(ref, 'Data', { startRow: 5, endRow: 5, startColumn: 2, endColumn: 4 }),
    ).toBe(false)
    expect(
      refIntersects(ref, 'Other', { startRow: 5, endRow: 5, startColumn: 1, endColumn: 1 }),
    ).toBe(false)
  })
})

describe('chartDataFromValues orientation and header detection', () => {
  it('wide cross-tab: rows become series, numeric year headers become categories', async () => {
    const { chartDataFromValues } = await import('../src/domain/chart-visual')
    const parsed = chartDataFromValues([
      ['', 2020, 2021, 2022],
      ['Division 1', 225, 210, 211.5],
      ['Division 2', 125, 157.5, 161.4],
    ])
    expect(parsed?.byRow).toBe(true)
    expect(parsed?.hasHeaderRow).toBe(true)
    expect(parsed?.hasCategoryColumn).toBe(true)
    expect(parsed?.categories).toEqual(['2020', '2021', '2022'])
    expect(parsed?.series.map((s) => s.name)).toEqual(['Division 1', 'Division 2'])
    expect(parsed?.series[0]?.values).toEqual([225, 210, 211.5])
    // series offsets are ROW offsets within the selection in by-row mode
    expect(parsed?.series.map((s) => s.column)).toEqual([1, 2])
  })

  it('wide month table: one series per salesperson, months as categories', async () => {
    const { chartDataFromValues } = await import('../src/domain/chart-visual')
    const parsed = chartDataFromValues([
      ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
      ['Ann', 1, 2, 3, 4, 5, 6],
      ['Bob', 2, 3, 4, 5, 6, 7],
      ['Cal', 3, 4, 5, 6, 7, 8],
      ['Dee', 4, 5, 6, 7, 8, 9],
      ['Eve', 5, 6, 7, 8, 9, 10],
    ])
    expect(parsed?.byRow).toBe(true)
    expect(parsed?.series.map((s) => s.name)).toEqual(['Ann', 'Bob', 'Cal', 'Dee', 'Eve'])
    expect(parsed?.categories).toEqual(['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'])
  })

  it('tall two-column data keeps the column orientation (pie shape)', async () => {
    const { chartDataFromValues } = await import('../src/domain/chart-visual')
    const parsed = chartDataFromValues([
      ['Apples', 10],
      ['Pears', 20],
      ['Plums', 30],
    ])
    expect(parsed?.byRow).toBe(false)
    expect(parsed?.series).toHaveLength(1)
    expect(parsed?.categories).toEqual(['Apples', 'Pears', 'Plums'])
  })

  it('keeps an all-numeric block header-less', async () => {
    const { chartDataFromValues } = await import('../src/domain/chart-visual')
    const parsed = chartDataFromValues([
      [1, 2],
      [3, 4],
    ])
    expect(parsed?.hasHeaderRow).toBe(false)
    expect(parsed?.series).toHaveLength(2)
  })

  it('charts a mixed first column when every other column is text', async () => {
    // Numbered checklists: the only numeric column doubles as the row-label
    // column, so claiming it for the category axis left zero series.
    const { chartDataFromValues } = await import('../src/domain/chart-visual')
    const parsed = chartDataFromValues([
      ['Monitoring checklist', null],
      [null, null],
      ['No.', 'Task'],
      [1, 'cart'],
      [2, 'cart'],
      [3, 'carousel'],
    ])
    expect(parsed?.hasCategoryColumn).toBe(false)
    expect(parsed?.series).toHaveLength(1)
    expect(parsed?.series[0]?.column).toBe(0)
    expect(parsed?.series[0]?.values).toEqual([0, 0, 0, 1, 2, 3])
  })

  it('charts a sparse numeric first column between blank filler rows', async () => {
    const { chartDataFromValues } = await import('../src/domain/chart-visual')
    const parsed = chartDataFromValues([
      [null, null],
      ['id', 'employee'],
      [1, 'Alice'],
      [2, 'Bob'],
      ['total', null],
      [null, null],
    ])
    expect(parsed?.series).toHaveLength(1)
    expect(parsed?.series[0]?.column).toBe(0)
  })

  it('keeps a mixed later column as a series despite text notes', async () => {
    const { chartDataFromValues } = await import('../src/domain/chart-visual')
    const parsed = chartDataFromValues([
      ['label', 'value'],
      ['a', 1],
      ['b', 'n/a'],
      ['c', 3],
    ])
    expect(parsed?.hasCategoryColumn).toBe(true)
    expect(parsed?.series[0]?.values).toEqual([1, 0, 3])
  })

  it('still rejects a range with no numeric cells anywhere', async () => {
    const { chartDataFromValues } = await import('../src/domain/chart-visual')
    const parsed = chartDataFromValues([
      ['question', 'answer'],
      ['agree', 'agree'],
      ['disagree', 'agree'],
    ])
    expect(parsed).toBeNull()
  })
})

describe('transposeChartSeries', () => {
  it('pivots categories into series and series names into categories', async () => {
    const { transposeChartSeries } = await import('../src/domain/chart-visual')
    const seriesSet = transposeChartSeries(
      [
        { name: 'Grinsley', categories: ['Jan', 'Feb'], values: [1, 2] },
        { name: 'Abbott', categories: ['Jan', 'Feb'], values: [3, 4] },
      ],
      (n) => `Series ${n}`,
    )
    expect(seriesSet).toEqual([
      { name: 'Jan', values: [1, 3], categories: ['Grinsley', 'Abbott'] },
      { name: 'Feb', values: [2, 4], categories: ['Grinsley', 'Abbott'] },
    ])
  })

  it('returns null when there are no categories to pivot on', async () => {
    const { transposeChartSeries } = await import('../src/domain/chart-visual')
    expect(
      transposeChartSeries([{ name: 'S1', categories: [], values: [1] }], (n) => `${n}`),
    ).toBeNull()
    expect(transposeChartSeries([], (n) => `${n}`)).toBeNull()
  })

  it('fills gaps with zeros and labels blank names', async () => {
    const { transposeChartSeries } = await import('../src/domain/chart-visual')
    const seriesSet = transposeChartSeries(
      [{ name: '', categories: ['', 'B'], values: [5] }],
      (n) => `Series ${n}`,
    )
    expect(seriesSet).toEqual([
      { name: 'Series 1', values: [5], categories: ['Series 1'] },
      { name: 'B', values: [0], categories: ['Series 1'] },
    ])
  })
})

describe('formatCategoryLabel', () => {
  it('formats numeric category text through its number format', () => {
    expect(formatCategoryLabel('44562', 'mmm\\-yy')).toBe('Jan-22')
    expect(formatCategoryLabel('41387', 'mmm\\ yyyy')).toBe('Apr 2013')
    expect(formatCategoryLabel('0.152', '0.0%')).toBe('15.2%')
  })

  it('leaves non-numeric, unformatted, and General labels alone', () => {
    expect(formatCategoryLabel('North', 'mmm\\-yy')).toBe('North')
    expect(formatCategoryLabel('44562', undefined)).toBe('44562')
    expect(formatCategoryLabel('44562', 'General')).toBe('44562')
    expect(formatCategoryLabel('', '0.0%')).toBe('')
  })
})

describe('chartCategoryFormat', () => {
  const chart = (
    axis: string | undefined,
    series: string | undefined,
  ): Pick<ChartVisualState, 'categoryAxisFormat' | 'series'> => ({
    categoryAxisFormat: axis,
    series: [{ name: 'S1', categories: ['44562'], values: [1], categoryFormat: series }],
  })

  it('prefers the axis-level format over the series numCache one', () => {
    expect(chartCategoryFormat(chart('d-mmm', 'mmm\\-yy'))).toBe('d-mmm')
    expect(chartCategoryFormat(chart(undefined, 'mmm\\-yy'))).toBe('mmm\\-yy')
    expect(chartCategoryFormat(chart('General', 'mmm\\-yy'))).toBe('mmm\\-yy')
    expect(chartCategoryFormat(chart(undefined, 'General'))).toBeUndefined()
    expect(chartCategoryFormat({ series: [] })).toBeUndefined()
  })
})

describe('scatterAxisBounds', () => {
  it('auto-scales positive data from 0 to an Excel-like nice ceiling', () => {
    expect(scatterAxisBounds([0.0026, 0.05, 0.152])).toMatchObject({ min: 0 })
    expect(scatterAxisBounds([0.0026, 0.05, 0.152]).max).toBeCloseTo(0.16, 10)
    expect(scatterAxisBounds([0.048, 0.2, 0.449]).max).toBeCloseTo(0.45, 10)
    expect(scatterAxisBounds([3, 40, 87]).max).toBe(90)
    expect(scatterAxisBounds([3, 40, 130]).max).toBe(140)
  })

  it('gives 5 evenly spaced ticks over the range', () => {
    const { ticks } = scatterAxisBounds([0.0026, 0.152])
    expect(ticks).toHaveLength(5)
    expect(ticks[0]).toBe(0)
    expect(ticks[1]).toBeCloseTo(0.04, 10)
    expect(ticks[4]).toBeCloseTo(0.16, 10)
  })

  it('honors explicit c:scaling bounds', () => {
    expect(scatterAxisBounds([2, 8], { min: 1, max: 10 })).toMatchObject({ min: 1, max: 10 })
    expect(scatterAxisBounds([2, 8], { max: 20 })).toMatchObject({ min: 0, max: 20 })
  })

  it('drops below zero only for negative data, and survives degenerate input', () => {
    const negative = scatterAxisBounds([-0.3, 0.5])
    expect(negative.min).toBeLessThanOrEqual(-0.3)
    expect(negative.max).toBeCloseTo(0.5, 10)
    expect(scatterAxisBounds([])).toMatchObject({ min: 0, max: 1 })
    expect(scatterAxisBounds([0, 0])).toMatchObject({ min: 0, max: 1 })
    expect(scatterAxisBounds([Number.NaN])).toMatchObject({ min: 0, max: 1 })
  })
})

describe('chartDataFromValues scatter X column', () => {
  it('routes a numeric first column into categories for scatter', async () => {
    const { chartDataFromValues } = await import('../src/domain/chart-visual')
    const parsed = chartDataFromValues(
      [
        ['Sales', 'EBIT'],
        [0.0626, 0.152],
        [0.021, 0.449],
      ],
      { numericCategoryColumn: true },
    )
    expect(parsed?.hasCategoryColumn).toBe(true)
    expect(parsed?.categories).toEqual(['0.0626', '0.021'])
    expect(parsed?.series).toHaveLength(1)
    expect(parsed?.series[0]).toMatchObject({ name: 'EBIT', values: [0.152, 0.449] })
  })

  it('by-row selections pivot the first data row into X', async () => {
    const { chartDataFromValues } = await import('../src/domain/chart-visual')
    // 2 rows × 4 cols with a label column: row 1 = X, row 2 = Y (corpus shape)
    const parsed = chartDataFromValues(
      [
        ['$ new shares', 100, 200, 300],
        ['Avg price', 10, 20, 30],
      ],
      { numericCategoryColumn: true },
    )
    expect(parsed?.byRow).toBe(true)
    expect(parsed?.hasCategoryColumn).toBe(true)
    expect(parsed?.categories).toEqual(['100', '200', '300'])
    expect(parsed?.series).toHaveLength(1)
    expect(parsed?.series[0]).toMatchObject({ name: 'Avg price', values: [10, 20, 30] })
  })

  it('non-scatter parsing is unchanged (each numeric column a series)', async () => {
    const { chartDataFromValues } = await import('../src/domain/chart-visual')
    const parsed = chartDataFromValues([
      ['Sales', 'EBIT'],
      [0.0626, 0.152],
      [0.021, 0.449],
    ])
    expect(parsed?.hasCategoryColumn).toBe(false)
    expect(parsed?.series).toHaveLength(2)
  })
})

describe('valueAxisScale', () => {
  it('matches Excel defaults on the run5 corpus', () => {
    // 60509: data max 11162 → 0..12000 step 2000
    expect(valueAxisScale(11162)).toEqual({
      min: 0,
      max: 12000,
      ticks: [0, 2000, 4000, 6000, 8000, 10000, 12000],
    })
    // budget: 3750 → 0..4000 step 500
    expect(valueAxisScale(3750).max).toBe(4000)
    expect(valueAxisScale(3750).ticks).toHaveLength(9)
    // 57362: 13 → 0..14 step 2
    expect(valueAxisScale(13)).toEqual({ min: 0, max: 14, ticks: [0, 2, 4, 6, 8, 10, 12, 14] })
    // prod_055: 877 → 0..1000 step 100 (the earlier 900 reading came from a
    // page-clipped ref; the plot area runs past the 900 gridline).
    expect(valueAxisScale(877).max).toBe(1000)
    expect(valueAxisScale(877).ticks).toHaveLength(11)
  })

  it('leaves Excel 5% auto-max headroom, calibrated on the real-run1 refs', () => {
    // aspose_sample1 pivot chart: 18 → 0..20 step 2
    expect(valueAxisScale(18)).toEqual({
      min: 0,
      max: 20,
      ticks: [0, 2, 4, 6, 8, 10, 12, 14, 16, 18, 20],
    })
    // aspose_sampleModifyLineChart: 148 → 0..160 step 20
    expect(valueAxisScale(148).max).toBe(160)
    expect(valueAxisScale(148).ticks[1]).toBe(20)
    // aspose_sampleGetFonts stacked: 289753.76 → 0..350000 step 50000
    expect(valueAxisScale(289753.76).max).toBe(350000)
    expect(valueAxisScale(289753.76).ticks[1]).toBe(50000)
    // aspose_sampleDisableTextWrappingForDataLabels: 1000 → 0..1200 step
    // 200 (the 5% bump makes 11 intervals of 100, pushing the unit to 200)
    expect(valueAxisScale(1000).max).toBe(1200)
    expect(valueAxisScale(1000).ticks[1]).toBe(200)
    // phpss_32readwriteLineChartNoPointMarkers1: 3490 → 0..4000 step 500
    expect(valueAxisScale(3490).max).toBe(4000)
    expect(valueAxisScale(3490).ticks[1]).toBe(500)
  })

  it('honours explicit bounds and unit', () => {
    expect(valueAxisScale(999, { min: -180, max: 180, majorUnit: 60 }).ticks).toEqual([
      -180, -120, -60, 0, 60, 120, 180,
    ])
  })

  it('survives degenerate spans', () => {
    expect(valueAxisScale(0).max).toBeGreaterThan(0)
  })
})
