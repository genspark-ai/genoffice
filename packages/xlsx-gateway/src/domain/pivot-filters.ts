/// Shared model for pivot filters: label filters (by row/column field member
/// label) and value filters (by a data field's aggregate). Creation (baking) and
/// refresh (layout rebuild) share the same matching semantics: members failing
/// the filter are marked hidden (pivotField item@h), and their source rows are
/// dropped from all aggregates (incl. subtotals/grand totals).

/// Label filter: member label equals/contains/begins-with the given text
/// (case-insensitive).
export interface PivotLabelFilter {
  readonly kind: 'label'
  /// The filtered row/column dimension field (cache-field index).
  readonly field: number
  readonly op: 'equal' | 'contains' | 'beginsWith'
  readonly value: string
}

/// Value filter: filters the field's members by their aggregate on dataField
/// (top N / greater than a threshold / closed interval). Aggregate first, then
/// filter.
export interface PivotValueFilter {
  readonly kind: 'value'
  /// The filtered row/column dimension field (cache-field index).
  readonly field: number
  /// The data field it measures (dataFields index, OOXML iMeasureFld).
  readonly dataField: number
  readonly op: 'top' | 'greaterThan' | 'between'
  /// top: keep the count members with the largest aggregates.
  readonly count?: number | undefined
  /// greaterThan: aggregate > from; between: from ≤ aggregate ≤ to.
  readonly from?: number | undefined
  readonly to?: number | undefined
}

export type PivotFilterDef = PivotLabelFilter | PivotValueFilter

/// Label matching (case-insensitive).
export function matchesLabelFilter(filter: PivotLabelFilter, memberLabel: string): boolean {
  const member = memberLabel.toLowerCase()
  const value = filter.value.toLowerCase()
  switch (filter.op) {
    case 'equal':
      return member === value
    case 'contains':
      return member.includes(value)
    case 'beginsWith':
      return member.startsWith(value)
  }
}

/// Value filtering: given each member's aggregate (null = the member has no
/// aggregatable data and is never kept), returns the set of kept member keys.
/// top takes the first count in descending value order (ties keep the input
/// order, matching member display order).
export function allowedByValueFilter<Key>(
  filter: PivotValueFilter,
  totals: readonly { readonly key: Key; readonly total: number | null }[],
): Set<Key> {
  const allowed = new Set<Key>()
  if (filter.op === 'top') {
    const ranked = totals
      .filter((entry): entry is { key: Key; total: number } => entry.total !== null)
      .map((entry, order) => ({ ...entry, order }))
      .sort((a, b) => b.total - a.total || a.order - b.order)
      .slice(0, Math.max(1, filter.count ?? 10))
    for (const entry of ranked) allowed.add(entry.key)
    return allowed
  }
  for (const entry of totals) {
    if (entry.total === null) continue
    if (filter.op === 'greaterThan') {
      if (entry.total > (filter.from ?? 0)) allowed.add(entry.key)
    } else if (
      entry.total >= (filter.from ?? Number.NEGATIVE_INFINITY) &&
      entry.total <= (filter.to ?? Number.POSITIVE_INFINITY)
    ) {
      allowed.add(entry.key)
    }
  }
  return allowed
}

/// Validates that a filter description deserialized from OOXML / IPC is within
/// the supported range.
export function isValidPivotFilter(filter: unknown): filter is PivotFilterDef {
  if (typeof filter !== 'object' || filter === null) return false
  const candidate = filter as Record<string, unknown>
  if (typeof candidate.field !== 'number') return false
  if (candidate.kind === 'label') {
    return (
      (candidate.op === 'equal' || candidate.op === 'contains' || candidate.op === 'beginsWith') &&
      typeof candidate.value === 'string'
    )
  }
  if (candidate.kind === 'value') {
    if (typeof candidate.dataField !== 'number') return false
    if (candidate.op === 'top') {
      return typeof candidate.count === 'number' && candidate.count >= 1
    }
    if (candidate.op === 'greaterThan') return typeof candidate.from === 'number'
    if (candidate.op === 'between') {
      return typeof candidate.from === 'number' && typeof candidate.to === 'number'
    }
    return false
  }
  return false
}
