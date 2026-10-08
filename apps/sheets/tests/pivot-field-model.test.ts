import { describe, expect, it } from 'vitest'

import {
  EMPTY_PIVOT_MODEL,
  addFieldToArea,
  defaultPivotModel,
  fieldAreas,
  modelToConfig,
  moveChip,
  removeChip,
  toggleField,
  validatePivotModel,
  type PivotField,
  type PivotLayoutModel,
} from '../src/renderer/pivot-field-model'

const FIELDS: PivotField[] = [
  { label: 'Region', colIndex: 0, numeric: false },
  { label: 'Product', colIndex: 1, numeric: false },
  { label: 'Amount', colIndex: 2, numeric: true },
  { label: 'Units', colIndex: 3, numeric: true },
]

describe('Excel default placement', () => {
  it('checks text fields into Rows and numeric fields into Values (Sum)', () => {
    let model = toggleField(EMPTY_PIVOT_MODEL, FIELDS, 0)
    expect(model.rows).toEqual([0])
    model = toggleField(model, FIELDS, 2)
    expect(model.values).toEqual([{ fieldIndex: 2, agg: 'sum' }])
    model = toggleField(model, FIELDS, 1)
    expect(model.rows).toEqual([0, 1])
  })

  it('unchecking removes the field from every area', () => {
    const model = toggleField(toggleField(EMPTY_PIVOT_MODEL, FIELDS, 2), FIELDS, 2)
    expect(model.values).toEqual([])
    expect(fieldAreas(model, 2)).toEqual([])
  })

  it('starts a new pivot with the first text field on Rows and the first number on Values', () => {
    const model = defaultPivotModel(FIELDS)
    expect(model.rows).toEqual([0])
    expect(model.values).toEqual([{ fieldIndex: 2, agg: 'sum' }])
    const allNumeric = defaultPivotModel([
      { label: 'A', colIndex: 0, numeric: true },
      { label: 'B', colIndex: 1, numeric: true },
    ])
    expect(allNumeric.rows).toEqual([0])
    expect(allNumeric.values).toEqual([{ fieldIndex: 1, agg: 'sum' }])
    // One column: nothing is left for Values, so the model is invalid rather
    // than one the pane could never keep valid.
    const single = defaultPivotModel([{ label: 'A', colIndex: 0, numeric: false }])
    expect(single.values).toEqual([])
    expect(validatePivotModel(single)).toBe('needValue')
  })

  it('a text field dropped on Values counts', () => {
    const model = addFieldToArea(EMPTY_PIVOT_MODEL, FIELDS, 1, 'values')
    expect(model.values).toEqual([{ fieldIndex: 1, agg: 'count' }])
  })
})

describe('area transitions', () => {
  const base: PivotLayoutModel = {
    ...EMPTY_PIVOT_MODEL,
    rows: [0, 1],
    values: [{ fieldIndex: 2, agg: 'sum' }],
    groupings: { 1: { kind: 'date', dateUnit: 'month' } },
  }

  it('reorders within an area for every insertion slot', () => {
    expect(
      moveChip(base, FIELDS, { area: 'rows', index: 0 }, { area: 'rows', index: 2 }).rows,
    ).toEqual([1, 0])
    expect(
      moveChip(base, FIELDS, { area: 'rows', index: 1 }, { area: 'rows', index: 0 }).rows,
    ).toEqual([1, 0])
    expect(
      moveChip(base, FIELDS, { area: 'rows', index: 0 }, { area: 'rows', index: 1 }).rows,
    ).toEqual([0, 1])
  })

  it('moves an axis chip to Columns, keeping its grouping', () => {
    const moved = moveChip(base, FIELDS, { area: 'rows', index: 1 }, { area: 'columns', index: 0 })
    expect(moved.rows).toEqual([0])
    expect(moved.columns).toEqual([1])
    expect(moved.groupings[1]).toEqual({ kind: 'date', dateUnit: 'month' })
  })

  it('moves an axis chip to Filters as (All), dropping grouping', () => {
    const moved = moveChip(base, FIELDS, { area: 'rows', index: 1 }, { area: 'filters', index: 0 })
    expect(moved.filters).toEqual([{ fieldIndex: 1, item: null }])
    expect(moved.groupings[1]).toBeUndefined()
  })

  it('moves a values chip to Rows and an axis chip to Values', () => {
    const toRows = moveChip(base, FIELDS, { area: 'values', index: 0 }, { area: 'rows', index: 0 })
    expect(toRows.rows).toEqual([2, 0, 1])
    expect(toRows.values).toEqual([])
    const toValues = moveChip(
      base,
      FIELDS,
      { area: 'rows', index: 0 },
      { area: 'values', index: 1 },
    )
    expect(toValues.rows).toEqual([1])
    expect(toValues.values).toEqual([
      { fieldIndex: 2, agg: 'sum' },
      { fieldIndex: 0, agg: 'count' },
    ])
  })

  it('a field lives in one axis area at a time, but Values may repeat it', () => {
    const twice = addFieldToArea(base, FIELDS, 2, 'values')
    expect(twice.values).toHaveLength(2)
    const toColumns = addFieldToArea(base, FIELDS, 0, 'columns')
    expect(toColumns.rows).toEqual([1])
    expect(toColumns.columns).toEqual([0])
  })

  it('drags a field already in the area to the marked slot', () => {
    const three: PivotLayoutModel = { ...base, rows: [0, 1, 3] }
    expect(
      moveChip(three, FIELDS, { area: 'list', fieldIndex: 0 }, { area: 'rows', index: 3 }).rows,
    ).toEqual([1, 3, 0])
    expect(
      moveChip(three, FIELDS, { area: 'list', fieldIndex: 3 }, { area: 'rows', index: 0 }).rows,
    ).toEqual([3, 0, 1])
  })

  it('drags from the field list into a slot', () => {
    const model = moveChip(
      base,
      FIELDS,
      { area: 'list', fieldIndex: 3 },
      { area: 'rows', index: 1 },
    )
    expect(model.rows).toEqual([0, 3, 1])
  })

  it('removes chips and their per-field settings', () => {
    const removed = removeChip(base, { area: 'rows', index: 1 })
    expect(removed.rows).toEqual([0])
    expect(removed.groupings[1]).toBeUndefined()
    expect(removeChip(base, { area: 'values', index: 0 }).values).toEqual([])
  })

  it('calculated fields never leave Values', () => {
    const calc: PivotLayoutModel = {
      ...base,
      values: [{ fieldIndex: -1, agg: 'sum', calcName: 'Profit', formula: 'Amount-Units' }],
    }
    expect(moveChip(calc, FIELDS, { area: 'values', index: 0 }, { area: 'rows', index: 0 })).toBe(
      calc,
    )
  })
})

describe('validation and config', () => {
  it('needs one row field and one value', () => {
    expect(validatePivotModel(EMPTY_PIVOT_MODEL)).toBe('needRow')
    expect(validatePivotModel({ ...EMPTY_PIVOT_MODEL, rows: [0] })).toBe('needValue')
    expect(
      validatePivotModel({
        ...EMPTY_PIVOT_MODEL,
        rows: [0],
        columns: [1],
        values: [
          { fieldIndex: 2, agg: 'sum' },
          { fieldIndex: 3, agg: 'sum' },
        ],
      }),
    ).toBe('colNeedsSingleValue')
    expect(
      validatePivotModel({
        ...EMPTY_PIVOT_MODEL,
        rows: [0],
        values: [
          { fieldIndex: 2, agg: 'sum', name: 'Total' },
          { fieldIndex: 3, agg: 'sum', name: 'total' },
        ],
      }),
    ).toBe('dupValueName')
  })

  it('maps the model to the add_pivot configuration with page fields and value filters', () => {
    const model: PivotLayoutModel = {
      filters: [{ fieldIndex: 1, item: 'A' }],
      columns: [],
      rows: [0],
      values: [{ fieldIndex: 2, agg: 'average', name: 'Avg', filter: { op: 'top', count: 3 } }],
      groupings: { 0: { kind: 'range', rangeStep: 10 }, 3: { kind: 'range', rangeStep: 5 } },
      labelFilters: {},
    }
    const config = modelToConfig(model, 'A1:D9', 'F1')
    expect(config.pageFields).toEqual([{ fieldIndex: 1, item: 'A' }])
    expect(config.valueFilters).toEqual([{ valueIndex: 0, rule: { op: 'top', count: 3 } }])
    expect(config.values).toEqual([{ fieldIndex: 2, agg: 'average', name: 'Avg' }])
    // Groupings of fields no longer on an axis are dropped.
    expect(config.groupings).toEqual([{ fieldIndex: 0, rule: { kind: 'range', rangeStep: 10 } }])
  })
})
