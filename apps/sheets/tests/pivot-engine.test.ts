import { describe, expect, it } from 'vitest'

import {
  parsePivotDefinition,
  setPivotRefreshOnLoad,
} from '@genoffice/xlsx-gateway/gateway/xlsx-pivot'
import {
  PivotRefreshError,
  growPivotDefinition,
  recomputePivotData,
} from '@genoffice/xlsx-gateway/domain/pivot-engine'

const PIVOT_XML =
  '<pivotTableDefinition name="PivotTable1" cacheId="1">' +
  '<location ref="E3:H7" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>' +
  '<pivotFields count="3">' +
  '<pivotField axis="axisRow" showAll="0"><items count="3">' +
  '<item x="0"/><item x="1"/><item t="default"/></items></pivotField>' +
  '<pivotField axis="axisCol" showAll="0"><items count="3">' +
  '<item x="0"/><item x="1"/><item t="default"/></items></pivotField>' +
  '<pivotField dataField="1" showAll="0"/>' +
  '</pivotFields>' +
  '<rowFields count="1"><field x="0"/></rowFields>' +
  '<rowItems count="3"><i><x/></i><i><x v="1"/></i><i t="grand"><x/></i></rowItems>' +
  '<colFields count="1"><field x="1"/></colFields>' +
  '<colItems count="3"><i><x/></i><i><x v="1"/></i><i t="grand"><x/></i></colItems>' +
  '<dataFields count="1"><dataField name="Sum of Sales" fld="2"/></dataFields>' +
  '</pivotTableDefinition>'

const CACHE_XML =
  '<pivotCacheDefinition>' +
  '<cacheSource type="worksheet"><worksheetSource ref="A1:C7" sheet="Data"/></cacheSource>' +
  '<cacheFields count="3">' +
  '<cacheField name="Region"><sharedItems count="2"><s v="East"/><s v="West"/></sharedItems></cacheField>' +
  '<cacheField name="Product"><sharedItems count="2"><s v="A"/><s v="B"/></sharedItems></cacheField>' +
  '<cacheField name="Sales"><sharedItems containsString="0" containsNumber="1"/></cacheField>' +
  '</cacheFields>' +
  '</pivotCacheDefinition>'

const SOURCE = [
  ['Region', 'Product', 'Sales'],
  ['East', 'A', 10],
  ['East', 'B', 20],
  ['West', 'A', 30],
  ['West', 'B', 40],
  ['East', 'A', 5],
  ['West', 'B', 1],
]

describe('parsePivotDefinition', () => {
  it('extracts location, fields, axes, layout lines, and source', () => {
    const definition = parsePivotDefinition(PIVOT_XML, CACHE_XML)
    expect(definition.outputRef).toBe('E3:H7')
    expect(definition.firstDataRow).toBe(2)
    expect(definition.firstDataCol).toBe(1)
    expect(definition.sourceSheet).toBe('Data')
    expect(definition.sourceRef).toBe('A1:C7')
    expect(definition.fields.map((field) => field.name)).toEqual(['Region', 'Product', 'Sales'])
    expect(definition.fields[0]!.sharedItems).toEqual(['East', 'West'])
    expect(definition.rowFields).toEqual([0])
    expect(definition.colFields).toEqual([1])
    expect(definition.rowLines).toHaveLength(3)
    expect(definition.rowLines[2]!.t).toBe('grand')
    expect(definition.dataFields).toEqual([{ name: 'Sum of Sales', field: 2, subtotal: 'sum' }])
    expect(definition.unsupported).toEqual([])
  })

  it('flags bad calculated fields and unsupported filter types', () => {
    // The formula references a nonexistent field (Revenue), so calculated fields fail closed.
    const withFormula = CACHE_XML.replace(
      '<cacheField name="Sales">',
      '<cacheField name="Sales" formula="Revenue*2">',
    )
    expect(parsePivotDefinition(PIVOT_XML, withFormula).unsupported.join()).toContain(
      'calculated field',
    )
    // Date filters and similar types are unsupported, so pivotFilters fail closed.
    const withFilters = PIVOT_XML.replace(
      '</pivotTableDefinition>',
      '<filters count="1"><filter fld="0" type="dateBetween"/></filters></pivotTableDefinition>',
    )
    expect(parsePivotDefinition(withFilters, CACHE_XML).unsupported.join()).toContain(
      'pivotFilters',
    )
  })
})

describe('setPivotRefreshOnLoad', () => {
  it('adds the attribute and replaces an existing value', () => {
    expect(
      setPivotRefreshOnLoad(
        '<pivotCacheDefinition r:id="rId1"><cacheFields/></pivotCacheDefinition>',
      ),
    ).toContain('<pivotCacheDefinition refreshOnLoad="1" r:id="rId1">')
    expect(
      setPivotRefreshOnLoad(
        '<pivotCacheDefinition refreshOnLoad="0"><cacheFields/></pivotCacheDefinition>',
      ),
    ).toContain('refreshOnLoad="1"')
  })
})

describe('recomputePivotData', () => {
  it('recomputes the data area with grand totals', () => {
    const definition = parsePivotDefinition(PIVOT_XML, CACHE_XML)
    const result = recomputePivotData(definition, SOURCE)
    expect(result.data).toEqual([
      [15, 20, 35],
      [30, 41, 71],
      [45, 61, 106],
    ])
  })

  it('fails closed when the source grows a new category', () => {
    const definition = parsePivotDefinition(PIVOT_XML, CACHE_XML)
    expect(() => recomputePivotData(definition, [...SOURCE, ['North', 'A', 9]])).toThrow(
      PivotRefreshError,
    )
    expect(() => recomputePivotData(definition, [...SOURCE, ['North', 'A', 9]])).toThrow(/North/)
  })

  it('fails closed when the source header moved', () => {
    const definition = parsePivotDefinition(PIVOT_XML, CACHE_XML)
    const moved = [['Area', 'Product', 'Sales'], ...SOURCE.slice(1)]
    expect(() => recomputePivotData(definition, moved)).toThrow(/Region/)
  })

  it('excludes hidden items from members and totals', () => {
    const hidden = PIVOT_XML.replace(
      '<pivotField axis="axisCol" showAll="0"><items count="3"><item x="0"/><item x="1"/>',
      '<pivotField axis="axisCol" showAll="0"><items count="3"><item x="0"/><item x="1" h="1"/>',
    )
    const definition = parsePivotDefinition(hidden, CACHE_XML)
    const result = recomputePivotData(definition, SOURCE)
    // Product "B" is filtered out everywhere, including grand totals.
    expect(result.data[0]).toEqual([15, 0, 15])
    expect(result.data[2]).toEqual([45, 0, 45])
  })

  it('applies a page-field selection', () => {
    const paged = PIVOT_XML.replace(
      '<rowFields count="1"><field x="0"/></rowFields>',
      '<rowFields count="1"><field x="0"/></rowFields>' +
        '<pageFields count="1"><pageField fld="1" item="0"/></pageFields>',
    )
      .replace(
        '<pivotField axis="axisCol" showAll="0">',
        '<pivotField axis="axisPage" showAll="0">',
      )
      .replace('<colFields count="1"><field x="1"/></colFields>', '')
      .replace(
        '<colItems count="3"><i><x/></i><i><x v="1"/></i><i t="grand"><x/></i></colItems>',
        '',
      )
    const definition = parsePivotDefinition(paged, CACHE_XML)
    const result = recomputePivotData(definition, SOURCE)
    // Only Product=A rows count: East 15, West 30, grand 45.
    expect(result.data).toEqual([[15], [30], [45]])
  })

  it('supports average and count aggregations', () => {
    const averaged = PIVOT_XML.replace(
      '<dataField name="Sum of Sales" fld="2"/>',
      '<dataField name="Average of Sales" fld="2" subtotal="average"/>',
    )
    const definition = parsePivotDefinition(averaged, CACHE_XML)
    const result = recomputePivotData(definition, SOURCE)
    expect(result.data[0]).toEqual([7.5, 20, 35 / 3])
  })
})

describe('showDataAs (show values as)', () => {
  // Raw aggregate matrix (rows = East/West/total, columns = A/B/total):
  //   [15, 20,  35]
  //   [30, 41,  71]
  //   [45, 61, 106]
  const withShowDataAs = (mode: string): string =>
    PIVOT_XML.replace(
      '<dataField name="Sum of Sales" fld="2"/>',
      `<dataField name="Sum of Sales" fld="2" showDataAs="${mode}"/>`,
    )

  it('parses showDataAs and treats "normal" as unset', () => {
    expect(
      parsePivotDefinition(withShowDataAs('percentOfRow'), CACHE_XML).dataFields[0]!.showDataAs,
    ).toBe('percentOfRow')
    expect(
      parsePivotDefinition(withShowDataAs('normal'), CACHE_XML).dataFields[0]!.showDataAs,
    ).toBeUndefined()
  })

  it('flags unsupported showDataAs modes', () => {
    const definition = parsePivotDefinition(withShowDataAs('runningTotal'), CACHE_XML)
    expect(definition.unsupported.join()).toContain('show values as')
    expect(definition.dataFields[0]!.showDataAs).toBeUndefined()
  })

  it('percentOfTotal divides every cell by the grand total', () => {
    const definition = parsePivotDefinition(withShowDataAs('percentOfTotal'), CACHE_XML)
    const result = recomputePivotData(definition, SOURCE)
    expect(result.data).toEqual([
      [15 / 106, 20 / 106, 35 / 106],
      [30 / 106, 41 / 106, 71 / 106],
      [45 / 106, 61 / 106, 1],
    ])
  })

  it('percentOfRow divides by the row total (totals row included)', () => {
    const definition = parsePivotDefinition(withShowDataAs('percentOfRow'), CACHE_XML)
    const result = recomputePivotData(definition, SOURCE)
    expect(result.data).toEqual([
      [15 / 35, 20 / 35, 1],
      [30 / 71, 41 / 71, 1],
      [45 / 106, 61 / 106, 1],
    ])
  })

  it('percentOfCol divides by the column total (grand column included)', () => {
    const definition = parsePivotDefinition(withShowDataAs('percentOfCol'), CACHE_XML)
    const result = recomputePivotData(definition, SOURCE)
    expect(result.data).toEqual([
      [15 / 45, 20 / 61, 35 / 106],
      [30 / 45, 41 / 61, 71 / 106],
      [1, 1, 1],
    ])
  })

  it('denominator follows visible rows after hidden items/filters (not the raw grand total)', () => {
    // With Product=B hidden, only column A remains visible: East=15 / West=30 / total 45.
    const hidden = withShowDataAs('percentOfCol').replace(
      '<pivotField axis="axisCol" showAll="0"><items count="3"><item x="0"/><item x="1"/>',
      '<pivotField axis="axisCol" showAll="0"><items count="3"><item x="0"/><item x="1" h="1"/>',
    )
    const definition = parsePivotDefinition(hidden, CACHE_XML)
    const result = recomputePivotData(definition, SOURCE)
    expect(result.data[0]).toEqual([15 / 45, null, 15 / 45])
    expect(result.data[2]).toEqual([1, null, 1])
  })
})

describe('growPivotDefinition (automatic layout growth)', () => {
  it('returns as-is when there are no new categories (same object, grown=false)', () => {
    const definition = parsePivotDefinition(PIVOT_XML, CACHE_XML)
    const growth = growPivotDefinition(definition, SOURCE)
    expect(growth.grown).toBe(false)
    expect(growth.definition).toBe(definition)
  })

  it('new row-axis category: appends sharedItems/fieldItems and inserts new rows before the grand total', () => {
    const definition = parsePivotDefinition(PIVOT_XML, CACHE_XML)
    const source = [...SOURCE, ['North', 'A', 9]]
    const growth = growPivotDefinition(definition, source)
    expect(growth.grown).toBe(true)
    expect(growth.definition.fields[0]!.sharedItems).toEqual(['East', 'West', 'North'])
    expect(growth.definition.fieldItems[0]).toHaveLength(3)
    // Row lines: East, West, North, grand total.
    expect(growth.definition.rowLines.map((line) => line.t)).toEqual([
      'data',
      'data',
      'data',
      'grand',
    ])
    const result = recomputePivotData(growth.definition, source)
    expect(result.data).toEqual([
      [15, 20, 35],
      [30, 41, 71],
      [9, 0, 9],
      [54, 61, 115],
    ])
  })

  it('new column-axis category: columns grow before the grand-total column', () => {
    const definition = parsePivotDefinition(PIVOT_XML, CACHE_XML)
    const source = [...SOURCE, ['East', 'C', 3]]
    const growth = growPivotDefinition(definition, source)
    expect(growth.grown).toBe(true)
    expect(growth.definition.rowLines).toBe(definition.rowLines) // row axis unchanged
    expect(growth.definition.colLines.map((line) => line.t)).toEqual([
      'data',
      'data',
      'data',
      'grand',
    ])
    const result = recomputePivotData(growth.definition, source)
    expect(result.data).toEqual([
      [15, 20, 3, 38],
      [30, 41, 0, 71],
      [45, 61, 3, 109],
    ])
  })

  it('new categories in a report filter field only append members, layout rows unchanged', () => {
    const paged = PIVOT_XML.replace(
      '<rowFields count="1"><field x="0"/></rowFields>',
      '<rowFields count="1"><field x="0"/></rowFields>' +
        '<pageFields count="1"><pageField fld="1" item="0"/></pageFields>',
    )
      .replace(
        '<pivotField axis="axisCol" showAll="0">',
        '<pivotField axis="axisPage" showAll="0">',
      )
      .replace('<colFields count="1"><field x="1"/></colFields>', '')
      .replace(
        '<colItems count="3"><i><x/></i><i><x v="1"/></i><i t="grand"><x/></i></colItems>',
        '',
      )
    const definition = parsePivotDefinition(paged, CACHE_XML)
    const source = [...SOURCE, ['East', 'C', 3]]
    const growth = growPivotDefinition(definition, source)
    expect(growth.grown).toBe(true)
    expect(growth.definition.fields[1]!.sharedItems).toEqual(['A', 'B', 'C'])
    expect(growth.definition.rowLines).toBe(definition.rowLines)
    // The filter still selects Product=A: rows for the new category C are filtered out, values unchanged.
    expect(recomputePivotData(growth.definition, source).data).toEqual([[15], [30], [45]])
  })

  it('renamed headers / narrowed data source still error (safety checks kept)', () => {
    const definition = parsePivotDefinition(PIVOT_XML, CACHE_XML)
    const renamed = [['Area', 'Product', 'Sales'], ...SOURCE.slice(1)]
    expect(() => growPivotDefinition(definition, renamed)).toThrow(/Region/)
    const narrow = SOURCE.map((row) => row.slice(0, 2))
    expect(() => growPivotDefinition(definition, narrow)).toThrow(/data source may have moved/)
  })
})

describe('growPivotDefinition multi-level rows', () => {
  // Tabular layout with two-level rows (Region > Product): East{A,B}+subtotal, West{B}+subtotal, grand total.
  const PIVOT2_XML =
    '<pivotTableDefinition name="Pivot2" cacheId="2">' +
    '<location ref="E1:G7" firstHeaderRow="1" firstDataRow="1" firstDataCol="2"/>' +
    '<pivotFields count="3">' +
    '<pivotField axis="axisRow" showAll="0"><items count="3">' +
    '<item x="0"/><item x="1"/><item t="default"/></items></pivotField>' +
    '<pivotField axis="axisRow" showAll="0"><items count="3">' +
    '<item x="0"/><item x="1"/><item t="default"/></items></pivotField>' +
    '<pivotField dataField="1" showAll="0"/>' +
    '</pivotFields>' +
    '<rowFields count="2"><field x="0"/><field x="1"/></rowFields>' +
    '<rowItems count="6"><i><x/><x/></i><i r="1"><x v="1"/></i><i t="default"><x/></i>' +
    '<i><x v="1"/><x v="1"/></i><i t="default"><x v="1"/></i><i t="grand"><x/></i></rowItems>' +
    '<colItems count="1"><i/></colItems>' +
    '<dataFields count="1"><dataField name="Sum of Sales" fld="2"/></dataFields>' +
    '</pivotTableDefinition>'
  const SOURCE2 = [
    ['Region', 'Product', 'Sales'],
    ['East', 'A', 10],
    ['East', 'B', 20],
    ['West', 'B', 40],
  ]

  it('new outer members and new combinations inserted by level, subtotals/grand totals preserved', () => {
    const definition = parsePivotDefinition(PIVOT2_XML, CACHE_XML)
    const source = [...SOURCE2, ['West', 'C', 7], ['North', 'A', 5]]
    const growth = growPivotDefinition(definition, source)
    expect(growth.grown).toBe(true)
    expect(growth.definition.fields[0]!.sharedItems).toEqual(['East', 'West', 'North'])
    expect(growth.definition.fields[1]!.sharedItems).toEqual(['A', 'B', 'C'])
    // East{A,B}+subtotal, West{B,C}+subtotal, North{A}+subtotal, grand total.
    expect(growth.definition.rowLines.map((line) => line.t)).toEqual([
      'data',
      'data',
      'default',
      'data',
      'data',
      'default',
      'data',
      'default',
      'grand',
    ])
    const result = recomputePivotData(growth.definition, source)
    expect(result.data).toEqual([[10], [20], [30], [40], [7], [47], [5], [5], [82]])
  })

  it('compact/outline layout (data rows do not pin every level) refuses to grow', () => {
    // Turn the first two-level data row into one that fixes only the outer level (an outline-style parent caption row).
    const outline = PIVOT2_XML.replace(
      '<rowItems count="6"><i><x/><x/></i>',
      '<rowItems count="6"><i><x/></i>',
    )
    const definition = parsePivotDefinition(outline, CACHE_XML)
    expect(() => growPivotDefinition(definition, [...SOURCE2, ['North', 'A', 5]])).toThrow(
      /Compact\/outline/,
    )
  })
})
