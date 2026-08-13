import { describe, expect, it } from 'vitest'

import { allowedByValueFilter, matchesLabelFilter } from '../src/domain/pivot-filters'
import { growPivotDefinition, recomputePivotData } from '../src/domain/pivot-engine'
import { parsePivotDefinition } from '../src/gateway/xlsx-pivot'

describe('pivot filter primitives', () => {
  it('matches labels case-insensitively', () => {
    expect(
      matchesLabelFilter({ kind: 'label', field: 0, op: 'equal', value: 'east' }, 'East'),
    ).toBe(true)
    expect(
      matchesLabelFilter({ kind: 'label', field: 0, op: 'contains', value: 'ast' }, 'East'),
    ).toBe(true)
    expect(
      matchesLabelFilter({ kind: 'label', field: 0, op: 'beginsWith', value: 'East' }, 'Northeast'),
    ).toBe(false)
  })

  it('keeps top-N / greaterThan / between members by aggregate', () => {
    const totals = [
      { key: 'A', total: 10 },
      { key: 'B', total: 30 },
      { key: 'C', total: 20 },
      { key: 'D', total: null },
    ]
    expect([
      ...allowedByValueFilter(
        { kind: 'value', field: 0, dataField: 0, op: 'top', count: 2 },
        totals,
      ),
    ]).toEqual(['B', 'C'])
    expect([
      ...allowedByValueFilter(
        { kind: 'value', field: 0, dataField: 0, op: 'greaterThan', from: 15 },
        totals,
      ),
    ]).toEqual(['B', 'C'])
    expect([
      ...allowedByValueFilter(
        { kind: 'value', field: 0, dataField: 0, op: 'between', from: 10, to: 20 },
        totals,
      ),
    ]).toEqual(['A', 'C'])
  })
})

/// Minimal definition: single row field (Region: East/West/South) + Sum of Sales; all three
/// members visible with full layout rows; filters live in <filters> and the engine re-applies them on refresh.
function pivotXmlWithFilters(filtersXml: string): string {
  return (
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
    filtersXml +
    '</pivotTableDefinition>'
  )
}

const FILTER_CACHE_XML =
  '<pivotCacheDefinition>' +
  '<cacheSource type="worksheet"><worksheetSource ref="A1:B7" sheet="Data"/></cacheSource>' +
  '<cacheFields count="2">' +
  '<cacheField name="Region"><sharedItems count="3">' +
  '<s v="East"/><s v="West"/><s v="South"/></sharedItems></cacheField>' +
  '<cacheField name="Sales"><sharedItems containsString="0" containsNumber="1"/></cacheField>' +
  '</cacheFields>' +
  '</pivotCacheDefinition>'

const FILTER_SOURCE = [
  ['Region', 'Sales'],
  ['East', 50],
  ['West', 30],
  ['South', 10],
]

describe('pivot filters recompute', () => {
  it('parses supported label and value filter types', () => {
    const definition = parsePivotDefinition(
      pivotXmlWithFilters(
        '<filters count="2">' +
          '<filter fld="0" type="captionBeginsWith" evalOrder="-1" id="1" stringValue1="E"/>' +
          '<filter fld="0" type="count" evalOrder="-1" id="2" iMeasureFld="0">' +
          '<autoFilter ref="A1"><filterColumn colId="0"><top10 val="2"/></filterColumn></autoFilter>' +
          '</filter>' +
          '</filters>',
      ),
      FILTER_CACHE_XML,
    )
    expect(definition.unsupported).toEqual([])
    expect(definition.filters).toEqual([
      { kind: 'label', field: 0, op: 'beginsWith', value: 'E' },
      { kind: 'value', field: 0, dataField: 0, op: 'top', count: 2 },
    ])
  })

  it('applies a label filter on refresh: non-matching members are hidden and dropped', () => {
    const definition = parsePivotDefinition(
      pivotXmlWithFilters(
        '<filters count="1">' +
          '<filter fld="0" type="captionEqual" evalOrder="-1" id="1" stringValue1="East"/>' +
          '</filters>',
      ),
      FILTER_CACHE_XML,
    )
    const growth = growPivotDefinition(definition, FILTER_SOURCE)
    expect(growth.grown).toBe(true)
    // West/South are filtered out and become hidden items; the layout keeps only East + grand total.
    expect(growth.definition.fieldItems[0]!.map((item) => item.hidden)).toEqual([false, true, true])
    const result = recomputePivotData(growth.definition, FILTER_SOURCE)
    expect(result.data).toEqual([[50], [50]])
  })

  it('applies a top-N value filter on refresh: aggregate first, then cut', () => {
    const definition = parsePivotDefinition(
      pivotXmlWithFilters(
        '<filters count="1">' +
          '<filter fld="0" type="count" evalOrder="-1" id="1" iMeasureFld="0">' +
          '<autoFilter ref="A1"><filterColumn colId="0"><top10 val="2"/></filterColumn></autoFilter>' +
          '</filter>' +
          '</filters>',
      ),
      FILTER_CACHE_XML,
    )
    const growth = growPivotDefinition(definition, FILTER_SOURCE)
    expect(growth.grown).toBe(true)
    // Aggregates East=50, West=30, South=10: the top 2 stay, South is hidden.
    const result = recomputePivotData(growth.definition, FILTER_SOURCE)
    expect(result.data).toEqual([[50], [30], [80]])
  })

  it('re-applies a value filter when the data changes the ranking', () => {
    const definition = parsePivotDefinition(
      pivotXmlWithFilters(
        '<filters count="1">' +
          '<filter fld="0" type="valueGreaterThan" evalOrder="-1" id="1" iMeasureFld="0">' +
          '<autoFilter ref="A1"><filterColumn colId="0"><customFilters>' +
          '<customFilter operator="greaterThan" val="25"/>' +
          '</customFilters></filterColumn></autoFilter>' +
          '</filter>' +
          '</filters>',
      ),
      FILTER_CACHE_XML,
    )
    // West drops below the threshold, so the second refresh filters it out too.
    const changed = [
      ['Region', 'Sales'],
      ['East', 50],
      ['West', 20],
      ['South', 10],
    ]
    const growth = growPivotDefinition(definition, changed)
    expect(growth.grown).toBe(true)
    const result = recomputePivotData(growth.definition, changed)
    expect(result.data).toEqual([[50], [50]])
  })

  it('does not rebuild when the hidden marks already match the filter', () => {
    // West/South are already hidden items (consistent with captionEqual=East), so no growth is needed.
    const alreadyHidden = pivotXmlWithFilters(
      '<filters count="1">' +
        '<filter fld="0" type="captionEqual" evalOrder="-1" id="1" stringValue1="East"/>' +
        '</filters>',
    )
      .replace(
        '<item x="0"/><item x="1"/><item x="2"/>',
        '<item x="0"/><item x="1" h="1"/><item x="2" h="1"/>',
      )
      .replace(
        '<rowItems count="4"><i><x/></i><i><x v="1"/></i><i><x v="2"/></i>',
        '<rowItems count="2"><i><x/></i>',
      )
    const definition = parsePivotDefinition(alreadyHidden, FILTER_CACHE_XML)
    const growth = growPivotDefinition(definition, FILTER_SOURCE)
    expect(growth.grown).toBe(false)
    expect(recomputePivotData(definition, FILTER_SOURCE).data).toEqual([[50], [50]])
  })
})
