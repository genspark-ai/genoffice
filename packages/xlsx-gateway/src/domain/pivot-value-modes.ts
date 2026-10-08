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
] as const
export type PivotAggregation = (typeof PIVOT_AGGREGATIONS)[number]

export const PIVOT_SHOW_DATA_AS = [
  'percentOfTotal',
  'percentOfRow',
  'percentOfCol',
  'percentOfParentRow',
  'percentOfParentCol',
  'index',
] as const
export type PivotShowDataAs = (typeof PIVOT_SHOW_DATA_AS)[number]

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
}

export function defaultDataFieldCaption(agg: PivotAggregation, field: string): string {
  return `${AGG_CAPTIONS[agg]} of ${field}`
}

/// Percent modes default to Excel's 0.00% (numFmtId 10); index stays General.
export function showDataAsIsPercent(mode: PivotShowDataAs | undefined): boolean {
  return mode !== undefined && mode !== 'index'
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
