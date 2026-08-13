import { describe, expect, it } from 'vitest'

import { applyPivotSlicer, PivotRefreshError, recomputePivotData } from '../src/domain/pivot-engine'
import { parsePivotDefinition } from '../src/gateway/xlsx-pivot'

/// Minimal definition: single row field (Region: East/West/South) + Sum of Sales (same as the
/// pivot-filters tests but without <filters>): the slicer drives hidden items directly.
const PIVOT_XML =
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

const CACHE_XML =
  '<pivotCacheDefinition>' +
  '<cacheSource type="worksheet"><worksheetSource ref="A1:B7" sheet="Data"/></cacheSource>' +
  '<cacheFields count="2">' +
  '<cacheField name="Region"><sharedItems count="3">' +
  '<s v="East"/><s v="West"/><s v="South"/></sharedItems></cacheField>' +
  '<cacheField name="Sales"><sharedItems containsString="0" containsNumber="1"/></cacheField>' +
  '</cacheFields>' +
  '</pivotCacheDefinition>'

const SOURCE = [
  ['Region', 'Sales'],
  ['East', 50],
  ['West', 30],
  ['South', 10],
]

describe('pivot slicer selection', () => {
  it('hides unselected members and drops their rows from layout and totals', () => {
    const definition = parsePivotDefinition(PIVOT_XML, CACHE_XML)
    const sliced = applyPivotSlicer(definition, SOURCE, 0, [0])
    // West/South are unselected and become hidden items; the layout keeps only East + grand total.
    expect(sliced.fieldItems[0]!.map((item) => item.hidden)).toEqual([false, true, true])
    expect(sliced.rowLines.map((line) => line.t)).toEqual(['data', 'grand'])
    const result = recomputePivotData(sliced, SOURCE)
    expect(result.data).toEqual([[50], [50]])
  })

  it('supports multi-select: every selected member stays, totals follow', () => {
    const definition = parsePivotDefinition(PIVOT_XML, CACHE_XML)
    const sliced = applyPivotSlicer(definition, SOURCE, 0, [1, 2])
    expect(sliced.fieldItems[0]!.map((item) => item.hidden)).toEqual([true, false, false])
    const result = recomputePivotData(sliced, SOURCE)
    // West=30, South=10, grand total 40 (East is excluded from all aggregates).
    expect(result.data).toEqual([[30], [10], [40]])
  })

  it('clearing the slicer (null selection) restores every member and the full layout', () => {
    const definition = parsePivotDefinition(PIVOT_XML, CACHE_XML)
    const sliced = applyPivotSlicer(definition, SOURCE, 0, [0])
    const restored = applyPivotSlicer(sliced, SOURCE, 0, null)
    expect(restored.fieldItems[0]!.map((item) => item.hidden)).toEqual([false, false, false])
    expect(restored.rowLines.map((line) => line.t)).toEqual(['data', 'data', 'data', 'grand'])
    const result = recomputePivotData(restored, SOURCE)
    expect(result.data).toEqual([[50], [30], [10], [90]])
  })

  it('is a no-op (same object) when the selection already matches visibility', () => {
    const definition = parsePivotDefinition(PIVOT_XML, CACHE_XML)
    expect(applyPivotSlicer(definition, SOURCE, 0, [0, 1, 2])).toBe(definition)
    expect(applyPivotSlicer(definition, SOURCE, 0, null)).toBe(definition)
  })

  it('rejects non-axis fields and unknown members (fail-closed)', () => {
    const definition = parsePivotDefinition(PIVOT_XML, CACHE_XML)
    // Field 1 (Sales) is a value field; it sits on no row/column/report-filter axis.
    expect(() => applyPivotSlicer(definition, SOURCE, 1, [0])).toThrow(PivotRefreshError)
    expect(() => applyPivotSlicer(definition, SOURCE, 0, [99])).toThrow(PivotRefreshError)
  })
})
