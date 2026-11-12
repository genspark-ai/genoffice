/// The add_chart `columns` parameter (#2528): named columns of the dataRange —
/// first the labels, then one series each — so non-adjacent columns chart
/// without copying data around, and the header/direction guesses are settled
/// by the caller instead of detected from the data.
import { describe, expect, it } from 'vitest'

import { InMemoryWorkbookAdapter } from '@genoffice/xlsx-gateway/domain/in-memory-workbook'
import { buildChartVisual } from '@genoffice/xlsx-gateway/domain/chart-visual'
import type { WorkbookSnapshot } from '@genoffice/xlsx-gateway/domain/workbook.types'

/** Item, Units, Price, Revenue — the issue's table: Revenue against Item
    crosses two columns a single dataRange block would drag along. */
function itemSnapshot(): WorkbookSnapshot {
  return {
    revision: 0,
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: {
          A1: { value: 'Item' },
          B1: { value: 'Units' },
          C1: { value: 'Price' },
          D1: { value: 'Revenue' },
          A2: { value: 'Widgets' },
          B2: { value: 10 },
          C2: { value: 2.5 },
          D2: { value: 25 },
          A3: { value: 'Gadgets' },
          B3: { value: 4 },
          C3: { value: 9 },
          D3: { value: 36 },
          A4: { value: 'Sprockets' },
          B4: { value: 7 },
          C4: { value: 3.25 },
          D4: { value: 22.75 },
        },
      },
    ],
  }
}

function itemGrid(): (string | number)[][] {
  return [
    ['Item', 'Units', 'Price', 'Revenue'],
    ['Widgets', 10, 2.5, 25],
    ['Gadgets', 4, 9, 36],
    ['Sprockets', 7, 3.25, 22.75],
  ]
}

describe('add_chart named columns', () => {
  it('charts two non-adjacent columns as labels + one series', () => {
    const visual = buildChartVisual({
      id: 'chart-1',
      sheetId: 'sheet-1',
      sheetName: 'Sheet1',
      chartType: 'column',
      dataRange: 'A1:D4',
      columns: ['A', 'D'],
      values: itemGrid(),
    })
    expect(visual.chart.series).toHaveLength(1)
    expect(visual.chart.series[0]).toMatchObject({
      name: 'Revenue',
      categories: ['Widgets', 'Gadgets', 'Sprockets'],
      values: [25, 36, 22.75],
    })
    // the series spans its own column only, the labels their named column
    expect(visual.chart.series[0]?.valuesRef).toBe("'Sheet1'!$D$2:$D$4")
    expect(visual.chart.series[0]?.categoriesRef).toBe("'Sheet1'!$A$2:$A$4")
    expect(visual.chart.title).toBe('Revenue')
  })

  it('charts the named columns in the given order, each as its own series', () => {
    const visual = buildChartVisual({
      id: 'chart-1',
      sheetId: 'sheet-1',
      sheetName: 'Sheet1',
      chartType: 'column',
      dataRange: 'A1:D4',
      columns: ['A', 'D', 'B'],
      values: itemGrid(),
    })
    expect(visual.chart.series.map((series) => series.name)).toEqual(['Revenue', 'Units'])
    expect(visual.chart.series[0]?.valuesRef).toBe("'Sheet1'!$D$2:$D$4")
    expect(visual.chart.series[1]?.valuesRef).toBe("'Sheet1'!$B$2:$B$4")
  })

  it('treats the first row as the header even when the headers are numeric', () => {
    // Numeric headers (years) are data to the detector; naming the columns
    // states that they are headers.
    const visual = buildChartVisual({
      id: 'chart-1',
      sheetId: 'sheet-1',
      sheetName: 'Sheet1',
      chartType: 'line',
      dataRange: 'A1:D4',
      columns: ['A', 'C', 'D'],
      values: [
        ['Item', null, 2023, 2024],
        ['Widgets', null, 1, 2],
        ['Gadgets', null, 3, 4],
        ['Sprockets', null, 5, 6],
      ],
    })
    expect(visual.chart.series.map((series) => series.name)).toEqual(['2023', '2024'])
    expect(visual.chart.series[0]?.values).toEqual([1, 3, 5])
  })

  it('accepts lowercase column letters', () => {
    const visual = buildChartVisual({
      id: 'chart-1',
      sheetId: 'sheet-1',
      sheetName: 'Sheet1',
      chartType: 'column',
      dataRange: 'A1:D4',
      columns: ['a', 'd'],
      values: itemGrid(),
    })
    expect(visual.chart.series[0]?.valuesRef).toBe("'Sheet1'!$D$2:$D$4")
  })

  it('refuses a named column with no numbers', () => {
    expect(() =>
      buildChartVisual({
        id: 'chart-1',
        sheetId: 'sheet-1',
        sheetName: 'Sheet1',
        chartType: 'column',
        dataRange: 'A1:D4',
        columns: ['B', 'A'],
        values: itemGrid(),
      }),
    ).toThrow(/A has no numbers to chart/)
  })

  it('refuses a column outside the dataRange, duplicates and single-column lists', () => {
    const base = {
      id: 'chart-1',
      sheetId: 'sheet-1',
      sheetName: 'Sheet1',
      chartType: 'column' as const,
      dataRange: 'A1:D4',
      values: itemGrid(),
    }
    expect(() => buildChartVisual({ ...base, columns: ['A', 'E'] })).toThrow(
      /E lies outside the dataRange/,
    )
    expect(() => buildChartVisual({ ...base, columns: ['A', 'D', 'D'] })).toThrow(/distinct/)
    expect(() => buildChartVisual({ ...base, columns: ['A'] })).toThrow(/label column/)
  })

  it('lands the named-column chart in the demo workbook through plan and apply', () => {
    const adapter = new InMemoryWorkbookAdapter(itemSnapshot())
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx-chart-columns',
      baseRevision: 0,
      summary: 'Add a chart of Item against Revenue',
      operations: [
        {
          op: 'add_chart',
          sheetId: 'sheet-1',
          chartType: 'column',
          dataRange: 'A1:D4',
          columns: ['A', 'D'],
        },
      ],
    })
    adapter.apply(plan)
    const visual = adapter.getSnapshot().sheets[0]?.visuals?.[0]
    expect(visual?.chart.series[0]).toMatchObject({
      name: 'Revenue',
      categories: ['Widgets', 'Gadgets', 'Sprockets'],
      values: [25, 36, 22.75],
    })
  })

  it('rejects malformed columns at plan time', () => {
    const adapter = new InMemoryWorkbookAdapter(itemSnapshot())
    const base = {
      dslVersion: 1 as const,
      baseRevision: 0,
      summary: 'Add a chart',
      operations: [
        {
          op: 'add_chart' as const,
          sheetId: 'sheet-1',
          chartType: 'column' as const,
          dataRange: 'A1:D4',
        },
      ],
    }
    expect(() =>
      adapter.plan({
        ...base,
        transactionId: 'tx-c1',
        operations: [{ ...base.operations[0]!, columns: ['A'] }],
      }),
    ).toThrow()
    expect(() =>
      adapter.plan({
        ...base,
        transactionId: 'tx-c2',
        operations: [{ ...base.operations[0]!, columns: ['A', 'not-a-column'] }],
      }),
    ).toThrow()
    expect(() =>
      adapter.plan({
        ...base,
        transactionId: 'tx-c3',
        operations: [{ ...base.operations[0]!, columns: Array.from({ length: 14 }, () => 'A') }],
      }),
    ).toThrow()
    // a column letter list of the right shape passes the schema, and the
    // plan's dry run already refuses a letter outside the range
    expect(() =>
      adapter.plan({
        ...base,
        transactionId: 'tx-c4',
        operations: [{ ...base.operations[0]!, columns: ['A', 'E'] }],
      }),
    ).toThrow(/E lies outside the dataRange/)
  })
})
