/**
 * add_chart's 2000-cell gate counts what the chart actually reads (genoffice#2835).
 * With `columns`, only the named columns are charted — the cells between them
 * never reach the visual — so a wide table with two named columns is cheap.
 */
import { describe, expect, it } from 'vitest'
import { parseRange, columnLabel } from '@genoffice/xlsx-gateway/domain/cell-address'
import { buildChartVisual, chartCellLimitError } from '@genoffice/xlsx-gateway/domain/chart-visual'
import { InMemoryWorkbookAdapter } from '@genoffice/xlsx-gateway/domain/in-memory-workbook'
import type { WorkbookSnapshot } from '@genoffice/xlsx-gateway/domain/workbook.types'

const ROWS = 251 // header + 250 data rows
const COLS = 10 // A..J, so the whole rectangle is 2510 cells

/** A = Date labels, J = Revenue numbers; the eight columns between are noise. */
function wideGrid(): (string | number)[][] {
  return Array.from({ length: ROWS }, (_, row) =>
    Array.from({ length: COLS }, (_, col) =>
      row === 0
        ? col === 0
          ? 'Date'
          : col === 9
            ? 'Revenue'
            : `Col${col}`
        : col === 0
          ? `Day ${row}`
          : col === 9
            ? row * 1.5
            : `noise-${row}-${col}`,
    ),
  )
}

function wideSnapshot(): WorkbookSnapshot {
  const cells: Record<string, { value: string | number }> = {}
  for (const [rowIndex, row] of wideGrid().entries()) {
    for (const [col, value] of row.entries()) {
      cells[`${columnLabel(col)}${rowIndex + 1}`] = { value }
    }
  }
  return {
    revision: 0,
    sheets: [{ id: 'sheet-1', name: 'Sheet1', cells }],
  }
}

const range = `A1:J${ROWS}`
const chart = (columns?: readonly string[]) => ({
  id: 'chart-1',
  sheetId: 'sheet-1',
  sheetName: 'Sheet1',
  chartType: 'line' as const,
  dataRange: range,
  ...(columns ? { columns } : {}),
  values: wideGrid(),
})

describe('chartCellLimitError', () => {
  it('counts only the named columns when they are given', () => {
    expect(chartCellLimitError(parseRange(range), ['A', 'J'])).toBeNull()
    expect(chartCellLimitError(parseRange(range))).toBe(
      'add_chart dataRange covers more than 2000 cells.',
    )
  })

  it('still refuses named columns that together exceed 2000 cells', () => {
    const bounds = parseRange('A1:J1001')
    // 2 * 1001 = 2002
    expect(chartCellLimitError(bounds, ['A', 'J'])).toBe(
      'add_chart columns cover more than 2000 cells.',
    )
    // 2 * 1000 = 2000 is the boundary and stays allowed
    expect(chartCellLimitError(parseRange('A1:J1000'), ['A', 'J'])).toBeNull()
  })
})

describe('add_chart with named columns on a wide range', () => {
  it('builds the visual from A and J alone', () => {
    const visual = buildChartVisual(chart(['A', 'J']))
    expect(visual.chart.series).toHaveLength(1)
    expect(visual.chart.series[0]).toMatchObject({ name: 'Revenue' })
    expect(visual.chart.series[0]?.values).toHaveLength(ROWS - 1)
    expect(visual.chart.series[0]?.valuesRef).toBe(`'Sheet1'!$J$2:$J$${ROWS}`)
  })

  it('keeps the rectangle gate when columns are absent', () => {
    expect(() => buildChartVisual(chart())).toThrow(/dataRange covers more than 2000 cells/)
  })

  it('passes the sheets plan and apply', () => {
    const adapter = new InMemoryWorkbookAdapter(wideSnapshot())
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx-wide-chart',
      baseRevision: 0,
      summary: 'Chart Revenue by Date',
      operations: [
        {
          op: 'add_chart',
          sheetId: 'sheet-1',
          chartType: 'line',
          dataRange: range,
          columns: ['A', 'J'],
        },
      ],
    })
    adapter.apply(plan)
    const applied = adapter.getSnapshot().sheets[0]?.visuals?.[0]
    expect(applied?.chart.series[0]).toMatchObject({ name: 'Revenue' })
    expect(applied?.chart.series[0]?.values).toHaveLength(ROWS - 1)
  })

  it('refuses the same wide range without columns at plan time', () => {
    const adapter = new InMemoryWorkbookAdapter(wideSnapshot())
    expect(() =>
      adapter.plan({
        dslVersion: 1,
        transactionId: 'tx-wide-chart-no-columns',
        baseRevision: 0,
        summary: 'Chart the whole range',
        operations: [
          {
            op: 'add_chart',
            sheetId: 'sheet-1',
            chartType: 'line',
            dataRange: range,
          },
        ],
      }),
    ).toThrow(/dataRange covers more than 2000 cells/)
  })
})
