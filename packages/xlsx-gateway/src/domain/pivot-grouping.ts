/// Pure-function model for pivot grouped fields (date/numeric ranges): groups a
/// dimension field's raw values into group labels by rule. Creation (baking the
/// grid + writing cache records) and recompute share the same label functions,
/// keeping the group-member space self-consistent and round-trippable in-engine.
///
/// TODO(OOXML persistence): ECMA-376 <fieldGroup> + <rangePr> + <groupItems> are
/// not generated yet — grouping params live in a private extension under
/// pivotTableDefinition/extLst (Excel ignores unknown ext, so the file remains
/// valid OOXML; a manual refresh in Excel degrades to grouping by raw values).
/// After reopening the file our engine reads the params back from extLst and can
/// still recompute with grouping.

export type PivotDateUnit = 'year' | 'quarter' | 'month'

export type PivotFieldGrouping =
  | { readonly kind: 'date'; readonly dateUnit: PivotDateUnit }
  | {
      readonly kind: 'range'
      /// Interval step (one bucket per rangeStep); must be positive.
      readonly rangeStep: number
      /// Interval start (default 0): bucket boundaries are rangeStart + k*rangeStep.
      readonly rangeStart?: number | undefined
    }

export type PivotGroupSourceValue = string | number | boolean | null | undefined

const MONTH_LABELS = [
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'May',
  'Jun',
  'Jul',
  'Aug',
  'Sep',
  'Oct',
  'Nov',
  'Dec',
]

/// Result of grouping one value: label is the group label (i.e. the grouped
/// dimension member; blank/unparseable values pass through verbatim as their own
/// group), sort is the in-group sort key (null for pass-through values, sorted
/// last).
export interface PivotGroupValue {
  readonly label: string
  readonly sort: number | null
}

/// Excel serial-date epoch: 1899-12-30 (compatible with the 1900 leap-year bug).
const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30)
const DAY_MS = 86_400_000

/// Parses a cell date: numbers as Excel serial dates, strings in YYYY-MM-DD /
/// YYYY/M/D read digits directly (avoiding Date.parse timezone ambiguity), and
/// everything else goes to Date.parse.
export function parseDateParts(value: string | number): { year: number; month: number } | null {
  if (typeof value === 'number') {
    if (!Number.isFinite(value) || value < 0 || value > 2_958_465) return null
    const date = new Date(EXCEL_EPOCH_UTC + Math.floor(value) * DAY_MS)
    return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1 }
  }
  const direct = /^\s*(\d{4})[-/](\d{1,2})(?:[-/](\d{1,2}))?/.exec(value)
  if (direct) {
    const year = Number(direct[1])
    const month = Number(direct[2])
    if (month >= 1 && month <= 12) return { year, month }
    return null
  }
  const parsed = Date.parse(value)
  if (!Number.isFinite(parsed)) return null
  const date = new Date(parsed)
  return { year: date.getFullYear(), month: date.getMonth() + 1 }
}

/// Collapses float bucket boundaries into stable short labels (removing noise
/// like 0.30000000000000004).
function formatBoundary(value: number): string {
  return String(Number(value.toFixed(10)))
}

/// Groups one raw value by the grouping rule. Blanks (null/undefined/'') return
/// an empty label (same blank semantics as ungrouped fields); non-blank values
/// that fail to parse pass through verbatim as labels, keeping problem data
/// visible instead of silently miscounting them into some bucket.
export function groupValue(
  grouping: PivotFieldGrouping,
  value: PivotGroupSourceValue,
): PivotGroupValue {
  if (value === null || value === undefined || value === '') {
    return { label: '', sort: null }
  }
  if (grouping.kind === 'date') {
    if (typeof value === 'boolean') return { label: String(value), sort: null }
    const parts = parseDateParts(value)
    if (!parts) return { label: String(value), sort: null }
    if (grouping.dateUnit === 'year') {
      return { label: String(parts.year), sort: parts.year }
    }
    if (grouping.dateUnit === 'quarter') {
      // Quarter/month units merge across years (single-level date grouping,
      // not nested under year).
      const quarter = Math.floor((parts.month - 1) / 3) + 1
      return { label: `Q${quarter}`, sort: quarter }
    }
    return { label: MONTH_LABELS[parts.month - 1]!, sort: parts.month }
  }
  const numeric = typeof value === 'number' ? value : Number(String(value).trim())
  if (!Number.isFinite(numeric)) return { label: String(value), sort: null }
  const start = grouping.rangeStart ?? 0
  const step = grouping.rangeStep
  const bucketStart = start + Math.floor((numeric - start) / step) * step
  // Labels are half-open intervals [bucketStart, bucketStart+step).
  return {
    label: `${formatBoundary(bucketStart)}-${formatBoundary(bucketStart + step)}`,
    sort: bucketStart,
  }
}

/// Label only (hot function on the recompute path).
export function groupLabel(grouping: PivotFieldGrouping, value: PivotGroupSourceValue): string {
  return groupValue(grouping, value).label
}

/// Validates that a grouping description deserialized from extLst / IPC is valid.
export function isValidGrouping(grouping: unknown): grouping is PivotFieldGrouping {
  if (typeof grouping !== 'object' || grouping === null) return false
  const candidate = grouping as {
    kind?: unknown
    dateUnit?: unknown
    rangeStep?: unknown
    rangeStart?: unknown
  }
  if (candidate.kind === 'date') {
    return (
      candidate.dateUnit === 'year' ||
      candidate.dateUnit === 'quarter' ||
      candidate.dateUnit === 'month'
    )
  }
  if (candidate.kind === 'range') {
    return (
      typeof candidate.rangeStep === 'number' &&
      Number.isFinite(candidate.rangeStep) &&
      candidate.rangeStep > 0 &&
      (candidate.rangeStart === undefined ||
        (typeof candidate.rangeStart === 'number' && Number.isFinite(candidate.rangeStart)))
    )
  }
  return false
}
