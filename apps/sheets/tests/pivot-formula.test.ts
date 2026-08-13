import { describe, expect, it } from 'vitest'

import {
  PivotFormulaError,
  evaluatePivotFormula,
  formatPivotFormula,
  parsePivotFormula,
} from '../src/domain/pivot-formula'
import { recomputePivotData } from '../src/domain/pivot-engine'
import { parsePivotDefinition } from '../src/gateway/xlsx-pivot'

const FIELDS = ['Revenue', 'Cost', 'Order Count']

function evaluate(formula: string, values: Record<string, number | null>): number | null {
  return evaluatePivotFormula(parsePivotFormula(formula, FIELDS), (name) => values[name] ?? null)
}

describe('pivot formula parser/evaluator', () => {
  it('handles arithmetic, precedence, parens, and unary minus', () => {
    expect(evaluate('Revenue-Cost', { Revenue: 100, Cost: 40 })).toBe(60)
    expect(evaluate('Revenue-Cost*2', { Revenue: 100, Cost: 40 })).toBe(20)
    expect(evaluate('(Revenue-Cost)/Revenue', { Revenue: 100, Cost: 40 })).toBe(0.6)
    expect(evaluate('-Cost+Revenue', { Revenue: 100, Cost: 40 })).toBe(60)
    expect(evaluate('Revenue*1.5+10', { Revenue: 100, Cost: 0 })).toBe(160)
  })

  it('supports quoted field names (with spaces) and a leading =', () => {
    expect(evaluate("='Order Count'*3", { 'Order Count': 7 })).toBe(21)
    // Case-insensitive field matching (canonicalized back to the original name).
    expect(
      evaluatePivotFormula(parsePivotFormula('revenue-cost', ['Revenue', 'Cost']), (name) =>
        name === 'Revenue' ? 10 : 4,
      ),
    ).toBe(6)
  })

  it('treats missing aggregates as 0 and division by zero as blank', () => {
    expect(evaluate('Revenue-Cost', { Revenue: 100, Cost: null })).toBe(100)
    expect(evaluate('Revenue/Cost', { Revenue: 100, Cost: 0 })).toBe(null)
  })

  it('rejects unknown fields, bad syntax, and unbalanced quotes', () => {
    expect(() => parsePivotFormula('Revenue-Margin', FIELDS)).toThrow(PivotFormulaError)
    expect(() => parsePivotFormula('Revenue+', FIELDS)).toThrow(PivotFormulaError)
    expect(() => parsePivotFormula('(Revenue-Cost', FIELDS)).toThrow(PivotFormulaError)
    expect(() => parsePivotFormula("'Revenue-Cost", FIELDS)).toThrow(PivotFormulaError)
    expect(() => parsePivotFormula('Revenue Cost', FIELDS)).toThrow(PivotFormulaError)
  })

  it('formats back to a quoted canonical form', () => {
    expect(formatPivotFormula(parsePivotFormula('Revenue-Cost*2', FIELDS))).toBe(
      "('Revenue'-('Cost'*2))",
    )
  })
})

/// Minimal definition: single row field + calculated field (profit = revenue - cost): the calculated
/// field is a cacheField@formula appended after the source fields, and the dataField's fld points at it.
const CALC_PIVOT_XML =
  '<pivotTableDefinition name="PivotTable1" cacheId="1">' +
  '<location ref="F1:H4" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>' +
  '<pivotFields count="4">' +
  '<pivotField axis="axisRow" showAll="0"><items count="3">' +
  '<item x="0"/><item x="1"/><item t="default"/></items></pivotField>' +
  '<pivotField dataField="1" showAll="0"/>' +
  '<pivotField dataField="1" showAll="0"/>' +
  '<pivotField dataField="1" showAll="0"/>' +
  '</pivotFields>' +
  '<rowFields count="1"><field x="0"/></rowFields>' +
  '<rowItems count="3"><i><x/></i><i><x v="1"/></i><i t="grand"><x/></i></rowItems>' +
  '<colFields count="1"><field x="-2"/></colFields>' +
  '<colItems count="2"><i><x/></i><i i="1"><x v="1"/></i></colItems>' +
  '<dataFields count="2">' +
  '<dataField name="Sum of Revenue" fld="1"/>' +
  '<dataField name="Sum of Profit" fld="3"/>' +
  '</dataFields>' +
  '</pivotTableDefinition>'

const CALC_CACHE_XML =
  '<pivotCacheDefinition>' +
  '<cacheSource type="worksheet"><worksheetSource ref="A1:C5" sheet="Data"/></cacheSource>' +
  '<cacheFields count="4">' +
  '<cacheField name="Region"><sharedItems count="2"><s v="East"/><s v="South"/></sharedItems></cacheField>' +
  '<cacheField name="Revenue"><sharedItems containsString="0" containsNumber="1"/></cacheField>' +
  '<cacheField name="Cost"><sharedItems containsString="0" containsNumber="1"/></cacheField>' +
  '<cacheField name="Profit" formula="\'Revenue\'-\'Cost\'" databaseField="0" numFmtId="0"/>' +
  '</cacheFields>' +
  '</pivotCacheDefinition>'

describe('calculated-field recompute', () => {
  it('parses the calculated field and evaluates it per group', () => {
    const definition = parsePivotDefinition(CALC_PIVOT_XML, CALC_CACHE_XML)
    expect(definition.unsupported).toEqual([])
    expect(definition.fields[3]!.formula).toBe("'Revenue'-'Cost'")
    expect(definition.dataFields[1]!.formula).toBe("'Revenue'-'Cost'")
    // The source has only 3 columns (the calculated field takes no source column); two value columns: revenue and profit.
    const result = recomputePivotData(definition, [
      ['Region', 'Revenue', 'Cost'],
      ['East', 100, 40],
      ['East', 50, 10],
      ['South', 80, 30],
    ])
    expect(result.data).toEqual([
      [150, 100], // East region: revenue 150, profit 150-50
      [80, 50], // South region: revenue 80, profit 80-30
      [230, 150], // grand total
    ])
  })

  it('flags a calculated field used as a dimension as unsupported', () => {
    const asAxis = CALC_PIVOT_XML.replace(
      '<rowFields count="1"><field x="0"/></rowFields>',
      '<rowFields count="1"><field x="3"/></rowFields>',
    )
    expect(parsePivotDefinition(asAxis, CALC_CACHE_XML).unsupported.join()).toContain(
      'used as a dimension',
    )
  })

  it('flags an unparseable calculated-field formula as unsupported', () => {
    const badFormula = CALC_CACHE_XML.replace(
      `formula="'Revenue'-'Cost'"`,
      'formula="Revenue-Margin"',
    )
    expect(parsePivotDefinition(CALC_PIVOT_XML, badFormula).unsupported.join()).toContain(
      'does not support recompute',
    )
  })
})
