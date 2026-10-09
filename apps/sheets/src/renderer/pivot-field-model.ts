/**
 * Pure layout model behind the PivotTable Fields pane: which source field sits
 * in which drop area (Filters / Columns / Rows / Values), how chips move
 * between and within areas, Excel's default placement rule, and the mapping
 * to the add_pivot configuration the app applies.
 */
import type {
  PivotAggregation,
  PivotShowDataAs,
} from '@genoffice/xlsx-gateway/domain/pivot-value-modes'

export interface PivotField {
  readonly label: string
  readonly colIndex: number
  /// Every non-blank source cell parses as a number (Excel's "numeric field").
  readonly numeric?: boolean | undefined
}

/// Grouping mode for dimension fields: dates by year/quarter/month, numbers by
/// fixed-step intervals.
export type PivotGroupingOption =
  | { readonly kind: 'date'; readonly dateUnit: 'year' | 'quarter' | 'month' }
  | { readonly kind: 'range'; readonly rangeStep: number }

export interface PivotLabelFilterOption {
  readonly op: 'equal' | 'contains' | 'beginsWith'
  readonly value: string
}

/// Value filter on a data field (top N / greater than / between), applied to
/// level-1 row-field members.
export interface PivotValueFilterOption {
  readonly op: 'top' | 'greaterThan' | 'between'
  readonly count?: number | undefined
  readonly from?: number | undefined
  readonly to?: number | undefined
}

export interface PivotValueSpec {
  /// Index into PivotField[]; -1 for calculated fields.
  readonly fieldIndex: number
  readonly agg: PivotAggregation
  readonly showDataAs?: PivotShowDataAs | undefined
  /// Custom caption; undefined = "Sum of <field>".
  readonly name?: string | undefined
  readonly calcName?: string | undefined
  readonly formula?: string | undefined
  readonly filter?: PivotValueFilterOption | undefined
}

export interface PivotPageSpec {
  readonly fieldIndex: number
  /// Selected member label; null = (All).
  readonly item: string | null
}

export type PivotArea = 'filters' | 'columns' | 'rows' | 'values'
export const PIVOT_AREAS: readonly PivotArea[] = ['filters', 'columns', 'rows', 'values']

export interface PivotLayoutModel {
  readonly filters: readonly PivotPageSpec[]
  readonly columns: readonly number[]
  readonly rows: readonly number[]
  readonly values: readonly PivotValueSpec[]
  readonly groupings: Readonly<Record<number, PivotGroupingOption>>
  readonly labelFilters: Readonly<Record<number, PivotLabelFilterOption>>
}

export interface PivotChipRef {
  readonly area: PivotArea
  readonly index: number
}

export type PivotDragSource = PivotChipRef | { readonly area: 'list'; readonly fieldIndex: number }

export const EMPTY_PIVOT_MODEL: PivotLayoutModel = {
  filters: [],
  columns: [],
  rows: [],
  values: [],
  groupings: {},
  labelFilters: {},
}

/// Full OOXML pivot configuration the app applies (create and relayout).
export interface OoXmlPivotConfig {
  readonly sourceRange: string
  readonly rowFieldIndices: readonly number[]
  readonly colFieldIndices: readonly number[]
  readonly pageFields: readonly PivotPageSpec[]
  readonly groupings: readonly { readonly fieldIndex: number; readonly rule: PivotGroupingOption }[]
  readonly labelFilters: readonly {
    readonly fieldIndex: number
    readonly rule: PivotLabelFilterOption
  }[]
  readonly valueFilters: readonly {
    readonly valueIndex: number
    readonly rule: PivotValueFilterOption
  }[]
  readonly values: readonly PivotValueSpec[]
  /// Top-left cell of the whole pivot block (report filter rows included).
  readonly targetCell: string
}

export type PivotModelError =
  | 'needRow'
  | 'needValue'
  | 'colNeedsSingleValue'
  | 'calcIncomplete'
  | 'rangeStep'
  | 'labelFilterText'
  | 'oneValueFilter'
  | 'filterConflict'
  | 'dupValueName'

function chipField(model: PivotLayoutModel, ref: PivotChipRef): number | undefined {
  switch (ref.area) {
    case 'filters':
      return model.filters[ref.index]?.fieldIndex
    case 'columns':
      return model.columns[ref.index]
    case 'rows':
      return model.rows[ref.index]
    case 'values':
      return model.values[ref.index]?.fieldIndex
  }
}

export function areaLength(model: PivotLayoutModel, area: PivotArea): number {
  return model[area].length
}

/// Areas a source field currently occupies (values may repeat a field).
export function fieldAreas(model: PivotLayoutModel, fieldIndex: number): PivotArea[] {
  const areas: PivotArea[] = []
  if (model.filters.some((page) => page.fieldIndex === fieldIndex)) areas.push('filters')
  if (model.columns.includes(fieldIndex)) areas.push('columns')
  if (model.rows.includes(fieldIndex)) areas.push('rows')
  if (model.values.some((value) => value.fieldIndex === fieldIndex)) areas.push('values')
  return areas
}

export function isFieldUsed(model: PivotLayoutModel, fieldIndex: number): boolean {
  return fieldAreas(model, fieldIndex).length > 0
}

/// Excel's checkbox rule: numeric fields land in Values (Sum), everything else
/// in Rows.
export function defaultAreaFor(field: PivotField): 'values' | 'rows' {
  return field.numeric === true ? 'values' : 'rows'
}

export function defaultAggFor(field: PivotField | undefined): PivotAggregation {
  return field?.numeric === true ? 'sum' : 'count'
}

/// Starting layout for a freshly created pivot: the first text field on Rows
/// and the first numeric field on Values (Excel starts empty, but the engine
/// needs one row field and one value to bake anything). A field never sits on
/// an axis and in Values at once, so a single-column source yields a model
/// that fails validation instead of one the pane cannot keep valid.
export function defaultPivotModel(fields: readonly PivotField[]): PivotLayoutModel {
  if (fields.length === 0) return EMPTY_PIVOT_MODEL
  const firstText = fields.findIndex((field) => field.numeric !== true)
  const rowField = firstText >= 0 ? firstText : 0
  const firstNumeric = fields.findIndex(
    (field, index) => field.numeric === true && index !== rowField,
  )
  const valueField = firstNumeric >= 0 ? firstNumeric : rowField === 0 ? 1 : 0
  if (valueField >= fields.length) return { ...EMPTY_PIVOT_MODEL, rows: [rowField] }
  return {
    ...EMPTY_PIVOT_MODEL,
    rows: [rowField],
    values: [{ fieldIndex: valueField, agg: defaultAggFor(fields[valueField]) }],
  }
}

function withoutField(model: PivotLayoutModel, fieldIndex: number): PivotLayoutModel {
  const groupings = { ...model.groupings }
  const labelFilters = { ...model.labelFilters }
  delete groupings[fieldIndex]
  delete labelFilters[fieldIndex]
  return {
    filters: model.filters.filter((page) => page.fieldIndex !== fieldIndex),
    columns: model.columns.filter((field) => field !== fieldIndex),
    rows: model.rows.filter((field) => field !== fieldIndex),
    values: model.values.filter((value) => value.fieldIndex !== fieldIndex),
    groupings,
    labelFilters,
  }
}

function insertAt<T>(list: readonly T[], index: number | undefined, item: T): T[] {
  const next = [...list]
  next.splice(Math.max(0, Math.min(index ?? next.length, next.length)), 0, item)
  return next
}

/// Adds a source field to an area (default: the end). A field lives in at most
/// one of Filters/Columns/Rows and never both on an axis and in Values, so the
/// other areas drop it first; Values may hold the same field several times.
export function addFieldToArea(
  model: PivotLayoutModel,
  fields: readonly PivotField[],
  fieldIndex: number,
  area: PivotArea,
  index?: number,
): PivotLayoutModel {
  if (fieldIndex < 0 && area !== 'values') return model
  if (area === 'values') {
    const base = fieldIndex < 0 ? model : withoutAxisField(model, fieldIndex)
    return {
      ...base,
      values: insertAt(base.values, index, {
        fieldIndex,
        agg: defaultAggFor(fields[fieldIndex]),
      }),
    }
  }
  // The slot is counted before the field's own chip (if any) leaves the area.
  const current =
    area === 'filters'
      ? model.filters.findIndex((page) => page.fieldIndex === fieldIndex)
      : model[area].indexOf(fieldIndex)
  const slot = index !== undefined && current >= 0 && current < index ? index - 1 : index
  const base = withoutField(model, fieldIndex)
  switch (area) {
    case 'filters':
      return { ...base, filters: insertAt(base.filters, slot, { fieldIndex, item: null }) }
    case 'columns':
      return { ...base, columns: insertAt(base.columns, slot, fieldIndex) }
    case 'rows':
      return { ...base, rows: insertAt(base.rows, slot, fieldIndex) }
  }
}

function withoutAxisField(model: PivotLayoutModel, fieldIndex: number): PivotLayoutModel {
  const stripped = withoutField(model, fieldIndex)
  return { ...stripped, values: model.values }
}

/// Checkbox toggle: unchecked → Excel's default area; checked → removed from
/// every area.
export function toggleField(
  model: PivotLayoutModel,
  fields: readonly PivotField[],
  fieldIndex: number,
): PivotLayoutModel {
  if (isFieldUsed(model, fieldIndex)) return withoutField(model, fieldIndex)
  const field = fields[fieldIndex]
  if (!field) return model
  return addFieldToArea(model, fields, fieldIndex, defaultAreaFor(field))
}

export function removeChip(model: PivotLayoutModel, ref: PivotChipRef): PivotLayoutModel {
  if (ref.area === 'values') {
    return { ...model, values: model.values.filter((_, index) => index !== ref.index) }
  }
  const fieldIndex = chipField(model, ref)
  if (fieldIndex === undefined) return model
  if (ref.area === 'filters') {
    return { ...model, filters: model.filters.filter((_, index) => index !== ref.index) }
  }
  const groupings = { ...model.groupings }
  const labelFilters = { ...model.labelFilters }
  delete groupings[fieldIndex]
  delete labelFilters[fieldIndex]
  const remaining = model[ref.area].filter((_, index) => index !== ref.index)
  return ref.area === 'rows'
    ? { ...model, rows: remaining, groupings, labelFilters }
    : { ...model, columns: remaining, groupings, labelFilters }
}

function reorder<T>(list: readonly T[], from: number, to: number): T[] {
  const next = [...list]
  const [item] = next.splice(from, 1)
  next.splice(Math.max(0, Math.min(to > from ? to - 1 : to, next.length)), 0, item!)
  return next
}

/// Drag-and-drop / menu move. `to.index` is the insertion slot in the target
/// area before the move (same-area moves adjust for the removed chip).
export function moveChip(
  model: PivotLayoutModel,
  fields: readonly PivotField[],
  from: PivotDragSource,
  to: { readonly area: PivotArea; readonly index: number },
): PivotLayoutModel {
  if (from.area === 'list') return addFieldToArea(model, fields, from.fieldIndex, to.area, to.index)
  const fieldIndex = chipField(model, from)
  if (fieldIndex === undefined) return model
  if (from.area === to.area) {
    switch (from.area) {
      case 'filters':
        return { ...model, filters: reorder(model.filters, from.index, to.index) }
      case 'columns':
        return { ...model, columns: reorder(model.columns, from.index, to.index) }
      case 'rows':
        return { ...model, rows: reorder(model.rows, from.index, to.index) }
      case 'values':
        return { ...model, values: reorder(model.values, from.index, to.index) }
    }
  }
  if (from.area === 'values') {
    if (fieldIndex < 0) return model
    const value = model.values[from.index]!
    const base: PivotLayoutModel = {
      ...model,
      values: model.values.filter((_, index) => index !== from.index),
    }
    // Keep sibling copies of the field in Values only when the move target is
    // Values itself; an axis placement excludes the field from Values.
    return addFieldToArea(base, fields, value.fieldIndex, to.area, to.index)
  }
  // Axis/filter chip → another area: a stripped copy of the model minus the
  // source chip, then the standard add (which also enforces uniqueness).
  const base = removeChip(model, from)
  const grouping = model.groupings[fieldIndex]
  const labelFilter = model.labelFilters[fieldIndex]
  const added = addFieldToArea(base, fields, fieldIndex, to.area, to.index)
  if (to.area === 'values' || to.area === 'filters') return added
  return {
    ...added,
    groupings: grouping ? { ...added.groupings, [fieldIndex]: grouping } : added.groupings,
    labelFilters: labelFilter
      ? { ...added.labelFilters, [fieldIndex]: labelFilter }
      : added.labelFilters,
  }
}

export function updateValueSpec(
  model: PivotLayoutModel,
  index: number,
  patch: Partial<PivotValueSpec>,
): PivotLayoutModel {
  return {
    ...model,
    values: model.values.map((value, at) => (at === index ? { ...value, ...patch } : value)),
  }
}

export function setPageItem(
  model: PivotLayoutModel,
  index: number,
  item: string | null,
): PivotLayoutModel {
  return {
    ...model,
    filters: model.filters.map((page, at) => (at === index ? { ...page, item } : page)),
  }
}

export function setFieldGrouping(
  model: PivotLayoutModel,
  fieldIndex: number,
  rule: PivotGroupingOption | null,
): PivotLayoutModel {
  const groupings = { ...model.groupings }
  if (rule === null) delete groupings[fieldIndex]
  else groupings[fieldIndex] = rule
  return { ...model, groupings }
}

export function setFieldLabelFilter(
  model: PivotLayoutModel,
  fieldIndex: number,
  rule: PivotLabelFilterOption | null,
): PivotLayoutModel {
  const labelFilters = { ...model.labelFilters }
  if (rule === null) delete labelFilters[fieldIndex]
  else labelFilters[fieldIndex] = rule
  return { ...model, labelFilters }
}

export function validatePivotModel(model: PivotLayoutModel): PivotModelError | null {
  if (model.rows.length === 0) return 'needRow'
  if (model.values.length === 0) return 'needValue'
  if (model.columns.length > 0 && model.values.length !== 1) return 'colNeedsSingleValue'
  if (
    model.values.some(
      (value) =>
        value.formula !== undefined &&
        ((value.calcName ?? '').trim() === '' || value.formula.trim() === ''),
    )
  ) {
    return 'calcIncomplete'
  }
  const names = model.values
    .map((value) => value.name?.trim().toLowerCase())
    .filter((name): name is string => name !== undefined && name !== '')
  if (new Set(names).size !== names.length) return 'dupValueName'
  const axisFields = new Set([...model.rows, ...model.columns])
  for (const [key, rule] of Object.entries(model.groupings)) {
    if (!axisFields.has(Number(key))) continue
    if (rule.kind === 'range' && (!Number.isFinite(rule.rangeStep) || rule.rangeStep <= 0)) {
      return 'rangeStep'
    }
  }
  for (const [key, rule] of Object.entries(model.labelFilters)) {
    if (axisFields.has(Number(key)) && rule.value.trim() === '') return 'labelFilterText'
  }
  const valueFilters = model.values.filter((value) => value.filter !== undefined)
  if (valueFilters.length > 1) return 'oneValueFilter'
  if (valueFilters.length === 1 && model.labelFilters[model.rows[0]!] !== undefined) {
    return 'filterConflict'
  }
  return null
}

export function modelToConfig(
  model: PivotLayoutModel,
  sourceRange: string,
  targetCell: string,
): OoXmlPivotConfig {
  const axisFields = new Set([...model.rows, ...model.columns])
  return {
    sourceRange,
    rowFieldIndices: [...model.rows],
    colFieldIndices: [...model.columns],
    pageFields: [...model.filters],
    groupings: Object.entries(model.groupings)
      .map(([key, rule]) => ({ fieldIndex: Number(key), rule }))
      .filter(({ fieldIndex }) => axisFields.has(fieldIndex)),
    labelFilters: Object.entries(model.labelFilters)
      .map(([key, rule]) => ({ fieldIndex: Number(key), rule }))
      .filter(({ fieldIndex }) => axisFields.has(fieldIndex)),
    valueFilters: model.values.flatMap((value, valueIndex) =>
      value.filter === undefined ? [] : [{ valueIndex, rule: value.filter }],
    ),
    values: model.values.map(({ filter: _filter, ...value }) => value),
    targetCell,
  }
}
