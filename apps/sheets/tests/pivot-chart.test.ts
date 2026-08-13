import { describe, expect, it } from 'vitest'

import { buildPivotChartData } from '../src/domain/pivot-chart'
import { parsePivotDefinition, type PivotDefinition } from '../src/gateway/xlsx-pivot'
import { createBufferEntrySource, planCellEditsToXlsx } from '../src/gateway/xlsx-gateway'
import type { SheetVisualAddition } from '../src/gateway/xlsx-gateway'
import { buildEditFixture } from './fixture-builder'

/// Minimal pivot definition: single row field (Region: East/West/South) + Sum of Sales,
/// output area E1:F5 (1 header row + 3 data rows + grand-total row).
const SINGLE_ROW_PIVOT_XML =
  '<pivotTableDefinition name="PivotTable1" cacheId="1">' +
  '<location ref="E1:F5" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>' +
  '<pivotFields count="2">' +
  '<pivotField axis="axisRow" showAll="0"><items count="4">' +
  '<item x="0"/><item x="1"/><item x="2"/><item t="default"/></items></pivotField>' +
  '<pivotField dataField="1" showAll="0"/>' +
  '</pivotFields>' +
  '<rowFields count="1"><field x="0"/></rowFields>' +
  '<rowItems count="4"><i><x/></i><i><x v="1"/></i><i><x v="2"/></i>' +
  '<i t="grand"><x/></i></rowItems>' +
  '<colItems count="1"><i/></colItems>' +
  '<dataFields count="1"><dataField name="Sum of Sales" fld="1"/></dataFields>' +
  '</pivotTableDefinition>'

const SINGLE_ROW_CACHE_XML =
  '<pivotCacheDefinition>' +
  '<cacheSource type="worksheet"><worksheetSource ref="A1:B7" sheet="Data"/></cacheSource>' +
  '<cacheFields count="2">' +
  '<cacheField name="Region"><sharedItems count="3">' +
  '<s v="East"/><s v="West"/><s v="South"/></sharedItems></cacheField>' +
  '<cacheField name="Sales"><sharedItems containsString="0" containsNumber="1"/></cacheField>' +
  '</cacheFields>' +
  '</pivotCacheDefinition>'

/// E1:F5 as a zero-based range.
const SINGLE_ROW_BOUNDS = { startRow: 0, startColumn: 4, endRow: 4, endColumn: 5 }

const SINGLE_ROW_OUTPUT: (string | number | null)[][] = [
  ['Region', 'Sum of Sales'],
  ['East', 50],
  ['West', 30],
  ['South', 10],
  ['Grand Total', 90],
]

describe('pivot chart data assembly', () => {
  it('maps the row axis to categories and the value column to a referenced series', () => {
    const definition = parsePivotDefinition(SINGLE_ROW_PIVOT_XML, SINGLE_ROW_CACHE_XML)
    const chart = buildPivotChartData(definition, 'Report', SINGLE_ROW_BOUNDS, SINGLE_ROW_OUTPUT)
    expect(chart.title).toBe('Sum of Sales')
    expect(chart.truncated).toBe(false)
    expect(chart.series).toHaveLength(1)
    const series = chart.series[0]!
    expect(series.name).toBe('Sum of Sales')
    // The grand-total row stays out of the chart: categories/values cover only the three data rows.
    expect(series.categories).toEqual(['East', 'West', 'South'])
    expect(series.values).toEqual([50, 30, 10])
    expect(series.valuesRef).toBe("'Report'!$F$2:$F$4")
    expect(series.categoriesRef).toBe("'Report'!$E$2:$E$4")
  })

  it('excludes grand-total columns and references each data column as a series', () => {
    // Region × Quarter one-row-one-column layout (hand-built definition focused on column-axis semantics):
    // output area A1:D5, 2 header rows, grand-total column at the end.
    const definition: PivotDefinition = {
      outputRef: 'A1:D5',
      firstDataRow: 2,
      firstDataCol: 1,
      fields: [
        { name: 'Region', sharedItems: ['East', 'West'] },
        { name: 'Quarter', sharedItems: ['Q1', 'Q2'] },
        { name: 'Sales', sharedItems: [] },
      ],
      fieldItems: [
        [
          { x: 0, hidden: false },
          { x: 1, hidden: false },
        ],
        [
          { x: 0, hidden: false },
          { x: 1, hidden: false },
        ],
        [],
      ],
      rowFields: [0],
      colFields: [1],
      rowLines: [
        { t: 'data', members: [0], depth: 1, dataField: 0 },
        { t: 'data', members: [1], depth: 1, dataField: 0 },
        { t: 'grand', members: [null], depth: 0, dataField: 0 },
      ],
      colLines: [
        { t: 'data', members: [0], depth: 1, dataField: 0 },
        { t: 'data', members: [1], depth: 1, dataField: 0 },
        { t: 'grand', members: [null], depth: 0, dataField: 0 },
      ],
      dataFields: [{ name: 'Sum of Sales', field: 2, subtotal: 'sum' }],
      pageFields: [],
      filters: [],
      sourceSheet: 'Data',
      sourceRef: 'A1:C9',
      unsupported: [],
    }
    const output: (string | number | null)[][] = [
      ['Sum of Sales', 'Quarter', null, null],
      ['Region', 'Q1', 'Q2', 'Grand Total'],
      ['East', 10, 20, 30],
      ['West', 5, 15, 20],
      ['Grand Total', 15, 35, 50],
    ]
    const chart = buildPivotChartData(
      definition,
      'Report',
      { startRow: 0, startColumn: 0, endRow: 4, endColumn: 3 },
      output,
    )
    // The grand-total column is excluded; each quarter gets a series, and the grand-total row stays out of categories.
    expect(chart.series.map((series) => series.name)).toEqual(['Q1', 'Q2'])
    expect(chart.series[0]!.values).toEqual([10, 5])
    expect(chart.series[1]!.values).toEqual([20, 15])
    expect(chart.series[0]!.valuesRef).toBe("'Report'!$B$3:$B$4")
    expect(chart.series[1]!.valuesRef).toBe("'Report'!$C$3:$C$4")
    expect(chart.series[0]!.categoriesRef).toBe("'Report'!$A$3:$A$4")
    expect(chart.series[0]!.categories).toEqual(['East', 'West'])
  })

  it('names one series per value field when the data-field axis is on columns', () => {
    const definition: PivotDefinition = {
      outputRef: 'E1:G4',
      firstDataRow: 1,
      firstDataCol: 1,
      fields: [
        { name: 'Region', sharedItems: ['East', 'West'] },
        { name: 'Sales', sharedItems: [] },
        { name: 'Qty', sharedItems: [] },
      ],
      fieldItems: [
        [
          { x: 0, hidden: false },
          { x: 1, hidden: false },
        ],
        [],
        [],
      ],
      rowFields: [0],
      colFields: [-2],
      rowLines: [
        { t: 'data', members: [0], depth: 1, dataField: 0 },
        { t: 'data', members: [1], depth: 1, dataField: 0 },
        { t: 'grand', members: [null], depth: 0, dataField: 0 },
      ],
      colLines: [
        { t: 'data', members: [0], depth: 1, dataField: 0 },
        { t: 'data', members: [1], depth: 1, dataField: 1 },
      ],
      dataFields: [
        { name: 'Sum of Sales', field: 1, subtotal: 'sum' },
        { name: 'Sum of Qty', field: 2, subtotal: 'sum' },
      ],
      pageFields: [],
      filters: [],
      sourceSheet: 'Data',
      sourceRef: 'A1:C9',
      unsupported: [],
    }
    const output: (string | number | null)[][] = [
      ['Region', 'Sum of Sales', 'Sum of Qty'],
      ['East', 50, 5],
      ['West', 30, 3],
      ['Grand Total', 80, 8],
    ]
    const chart = buildPivotChartData(
      definition,
      'Report',
      { startRow: 0, startColumn: 4, endRow: 3, endColumn: 6 },
      output,
    )
    expect(chart.title).toBe('PivotChart')
    expect(chart.series.map((series) => series.name)).toEqual(['Sum of Sales', 'Sum of Qty'])
    expect(chart.series[1]!.values).toEqual([5, 3])
    expect(chart.series[1]!.valuesRef).toBe("'Report'!$G$2:$G$3")
  })

  it('falls back to literal data when subtotal rows break the data-row run', () => {
    // Two-level rows (Region > Product) with the East subtotal row wedged between data rows: this
    // cannot be expressed as one contiguous range ref, so categories join the member captions per level.
    const definition: PivotDefinition = {
      outputRef: 'E1:H6',
      firstDataRow: 1,
      firstDataCol: 2,
      fields: [
        { name: 'Region', sharedItems: ['East', 'West'] },
        { name: 'Product', sharedItems: ['A'] },
        { name: 'Sales', sharedItems: [] },
      ],
      fieldItems: [
        [
          { x: 0, hidden: false },
          { x: 1, hidden: false },
        ],
        [{ x: 0, hidden: false }],
        [],
      ],
      rowFields: [0, 1],
      colFields: [],
      rowLines: [
        { t: 'data', members: [0, 0], depth: 2, dataField: 0 },
        { t: 'default', members: [0, null], depth: 1, dataField: 0 },
        { t: 'data', members: [1, 0], depth: 2, dataField: 0 },
        { t: 'grand', members: [null, null], depth: 0, dataField: 0 },
      ],
      colLines: [{ t: 'data', members: [], depth: 0, dataField: 0 }],
      dataFields: [{ name: 'Sum of Sales', field: 2, subtotal: 'sum' }],
      pageFields: [],
      filters: [],
      sourceSheet: 'Data',
      sourceRef: 'A1:C9',
      unsupported: [],
    }
    const output: (string | number | null)[][] = [
      ['Region', 'Product', 'Sum of Sales'],
      ['East', 'A', 50],
      ['East 小计', null, 50],
      ['West', 'A', 30],
      ['Grand Total', null, 80],
    ]
    const chart = buildPivotChartData(
      definition,
      'Report',
      { startRow: 0, startColumn: 4, endRow: 4, endColumn: 6 },
      output,
    )
    expect(chart.series).toHaveLength(1)
    const series = chart.series[0]!
    expect(series.categories).toEqual(['East - A', 'West - A'])
    expect(series.values).toEqual([50, 30])
    expect(series.valuesRef).toBeUndefined()
    expect(series.categoriesRef).toBeUndefined()
  })
})

describe('pivot chart persistence through visual additions', () => {
  it('writes a chart part whose series/cat reference the pivot output range', async () => {
    const definition = parsePivotDefinition(SINGLE_ROW_PIVOT_XML, SINGLE_ROW_CACHE_XML)
    const chart = buildPivotChartData(definition, 'Data', SINGLE_ROW_BOUNDS, SINGLE_ROW_OUTPUT)
    const addition: SheetVisualAddition = {
      sheetName: 'Data',
      anchor: {
        fromRow: 0,
        fromColumn: 7,
        fromRowOffset: 0,
        fromColumnOffset: 0,
        toRow: 15,
        toColumn: 14,
        toRowOffset: 0,
        toColumnOffset: 0,
      },
      chart: {
        chartType: 'column',
        title: chart.title,
        series: chart.series.map((series) => ({
          name: series.name,
          categories: [...series.categories],
          values: [...series.values],
          ...(series.valuesRef === undefined ? {} : { valuesRef: series.valuesRef }),
          ...(series.categoriesRef === undefined ? {} : { categoriesRef: series.categoriesRef }),
        })),
      },
    }
    const source = await createBufferEntrySource(await buildEditFixture())
    const plan = await planCellEditsToXlsx(
      source,
      [],
      [],
      [],
      undefined,
      [],
      [],
      [],
      [],
      [],
      null,
      [addition],
    )

    // The fixture already ships xl/charts/chart1.xml, so the new chart gets chart2.
    const chartXml = plan.added.get('xl/charts/chart2.xml')
    expect(chartXml).toBeDefined()
    expect(chartXml).toContain("<c:f>'Data'!$F$2:$F$4</c:f>")
    expect(chartXml).toContain("<c:f>'Data'!$E$2:$E$4</c:f>")
    expect(chartXml).toContain('<c:pt idx="0"><c:v>East</c:v></c:pt>')
    expect(chartXml).toContain('<c:pt idx="0"><c:v>50</c:v></c:pt>')
    // The grand-total row stays out of the cached points.
    expect(chartXml).not.toContain('<c:v>90</c:v>')

    const drawingXml = plan.added.get('xl/drawings/drawing1.xml')
    expect(drawingXml).toContain('<xdr:twoCellAnchor>')
    const drawingRels = plan.added.get('xl/drawings/_rels/drawing1.xml.rels')
    expect(drawingRels).toContain('Target="../charts/chart2.xml"')
    const contentTypes = plan.replaced.get('[Content_Types].xml')
    expect(contentTypes).toContain('PartName="/xl/charts/chart2.xml"')
  })
})
