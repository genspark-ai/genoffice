/// Pure-function model for pivot timelines: a month-granularity date-range
/// filter over one pivot dimension field. The range maps to a member set and
/// rides the same hidden-items mechanism as slicers (applyPivotSlicer).
import { parseDateParts } from './pivot-grouping'

/// Months count monotonically from year 0 (`year * 12 + month - 1`), so ranges
/// compare and iterate as plain integers across year boundaries.
export type MonthKey = number

export interface TimelineMember {
  /// fieldItems index (the member's entry position in the pivot cache).
  readonly member: number
  readonly monthKey: MonthKey
}

export interface TimelineDomain {
  readonly members: readonly TimelineMember[]
  readonly minMonth: MonthKey
  readonly maxMonth: MonthKey
}

export function monthKeyOf(value: string | number): MonthKey | null {
  const parts = parseDateParts(value)
  return parts === null ? null : parts.year * 12 + parts.month - 1
}

export function monthKeyParts(key: MonthKey): { year: number; month: number } {
  return { year: Math.floor(key / 12), month: (key % 12) + 1 }
}

/// Builds the timeline domain for one cache field. Returns null when the field
/// does not qualify as a date field: every non-blank shared item must parse as
/// a date (blank members stay out of the domain — a range hides them, clearing
/// the timeline shows them, matching Excel).
export function timelineDomainOf(
  sharedItems: readonly (string | number | boolean | null)[],
  fieldItems: readonly { readonly x: number | null; readonly hidden: boolean }[],
): TimelineDomain | null {
  const members: TimelineMember[] = []
  let min = Number.POSITIVE_INFINITY
  let max = Number.NEGATIVE_INFINITY
  for (const [member, item] of fieldItems.entries()) {
    if (item.x === null) continue
    const shared = sharedItems[item.x]
    if (shared === null || shared === undefined || shared === '') continue
    if (typeof shared === 'boolean') return null
    const key = monthKeyOf(shared)
    if (key === null) return null
    members.push({ member, monthKey: key })
    min = Math.min(min, key)
    max = Math.max(max, key)
  }
  if (members.length === 0) return null
  return { members, minMonth: min, maxMonth: max }
}

/// Member indices selected by an inclusive month range; null range = no filter.
export function timelineSelection(
  domain: TimelineDomain,
  range: { readonly start: MonthKey; readonly end: MonthKey } | null,
): readonly number[] | null {
  if (range === null) return null
  return domain.members
    .filter((entry) => entry.monthKey >= range.start && entry.monthKey <= range.end)
    .map((entry) => entry.member)
}
