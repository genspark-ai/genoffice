/// Data-field aggregation and "show values as" modes the pivot engine can bake,
/// recompute, and persist. Kept in one place so the DSL, the IPC schema, the
/// layout builder, the OOXML writer, and the parser agree on the envelope.

export const PIVOT_AGGREGATIONS = [
  'sum',
  'count',
  'average',
  'max',
  'min',
  'product',
  'countNums',
  'stdDev',
  'stdDevp',
  'var',
  'varp',
] as const
export type PivotAggregation = (typeof PIVOT_AGGREGATIONS)[number]

export const PIVOT_SHOW_DATA_AS = [
  'percentOfTotal',
  'percentOfRow',
  'percentOfCol',
  'percentOfParentRow',
  'percentOfParentCol',
  'index',
  'percentOf',
  'percentOfParent',
  'difference',
  'percentDiff',
  'runTotal',
  'percentOfRunningTotal',
  'rankAscending',
  'rankDescending',
] as const
export type PivotShowDataAs = (typeof PIVOT_SHOW_DATA_AS)[number]

/// Modes that reference a base field (the "Base field" picker of Excel's Show
/// Value As dialog): everything except the pure-ratio and index modes.
export const PIVOT_SHOW_DATA_AS_BASE_FIELD: readonly PivotShowDataAs[] = [
  'percentOf',
  'percentOfParent',
  'difference',
  'percentDiff',
  'runTotal',
  'percentOfRunningTotal',
  'rankAscending',
  'rankDescending',
]

/// Modes that also reference a base item of the base field (the "Base item"
/// picker): the difference/percent family. rank/runTotal/parent modes keep
/// baseItem at 0.
export const PIVOT_SHOW_DATA_AS_BASE_ITEM: readonly PivotShowDataAs[] = [
  'percentOf',
  'difference',
  'percentDiff',
]

export function showDataAsNeedsBaseField(mode: PivotShowDataAs): boolean {
  return PIVOT_SHOW_DATA_AS_BASE_FIELD.includes(mode)
}

export function showDataAsNeedsBaseItem(mode: PivotShowDataAs): boolean {
  return PIVOT_SHOW_DATA_AS_BASE_ITEM.includes(mode)
}

/// OOXML dataField@showDataAs spelling: our "% Of…" mode is plain "percent" in
/// the file (ECMA-376 ST_ShowDataAs); every other mode writes its own name.
export function showDataAsToOoxml(mode: PivotShowDataAs): string {
  return mode === 'percentOf' ? 'percent' : mode
}

/// Inverse of showDataAsToOoxml; null = not a supported showDataAs value.
export function showDataAsFromOoxml(value: string): PivotShowDataAs | null {
  if (value === 'percent') return 'percentOf'
  return isPivotShowDataAs(value) ? value : null
}

/// Excel's special dataField@baseItem values for "(previous)" / "(next)"
/// (MS-XLSB iBaseItem, reused by OOXML): an index larger than any item list.
/// Recognized only to fail closed on files that carry them.
export const PIVOT_BASE_ITEM_PREVIOUS = 1048828
export const PIVOT_BASE_ITEM_NEXT = 1048829

/// Excel's caption prefix per aggregation ("Sum of Amount"); Count Numbers
/// shares "Count of" with Count, as Excel does.
export const AGG_CAPTIONS: Record<PivotAggregation, string> = {
  sum: 'Sum',
  count: 'Count',
  average: 'Average',
  max: 'Max',
  min: 'Min',
  product: 'Product',
  countNums: 'Count',
  stdDev: 'StdDev',
  stdDevp: 'StdDevp',
  var: 'Var',
  varp: 'Varp',
}

export function defaultDataFieldCaption(agg: PivotAggregation, field: string): string {
  return `${AGG_CAPTIONS[agg]} of ${field}`
}

/// Modes that keep the General number format (no implicit 0.00%).
const NON_PERCENT_MODES: readonly string[] = [
  'difference',
  'runTotal',
  'rankAscending',
  'rankDescending',
  'index',
]

/// Percent modes default to Excel's 0.00% (numFmtId 10); difference, running
/// total, ranks and index stay General.
export function showDataAsIsPercent(mode: PivotShowDataAs | undefined): boolean {
  return mode !== undefined && !NON_PERCENT_MODES.includes(mode)
}

export function isPivotAggregation(value: string): value is PivotAggregation {
  return (PIVOT_AGGREGATIONS as readonly string[]).includes(value)
}

export function isPivotShowDataAs(value: string): value is PivotShowDataAs {
  return (PIVOT_SHOW_DATA_AS as readonly string[]).includes(value)
}

export interface PivotPageFieldSpec {
  readonly field: string
  readonly item?: string | undefined
}

/// add_pivot pageFields accept a bare header or { field, item }.
export function normalizePivotPageFields(
  pageFields: readonly (string | PivotPageFieldSpec)[] | undefined,
): PivotPageFieldSpec[] {
  return (pageFields ?? []).map((entry) =>
    typeof entry === 'string' ? { field: entry } : { field: entry.field, item: entry.item },
  )
}
