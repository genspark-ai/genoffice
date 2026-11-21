/** which records Finish & Merge takes: Word's Merge Range, with its From and To boxes */
export interface MergeRange {
  kind: 'all' | 'current' | 'custom'
  /** the first record of a custom range, counted from 1 */
  from: number
  /** the last record of a custom range; null runs to the end of the list */
  to: number | null
}

export const MERGE_RANGE_ALL: MergeRange = { kind: 'all', from: 1, to: null }

/** the records of the list the range takes; `current` is the index of the record the preview shows */
export function recordsInRange<T>(
  records: readonly T[],
  range: MergeRange,
  current: number,
): readonly T[] {
  if (range.kind === 'all') return records
  if (range.kind === 'current') return records.slice(current, current + 1)
  const from = Math.max(Math.trunc(range.from) || 1, 1)
  const to = Math.min(range.to === null ? records.length : Math.trunc(range.to), records.length)
  return from > to ? [] : records.slice(from - 1, to)
}

/** the place in the list of the first record the range takes, counted from 1 */
export function firstRecordInRange(range: MergeRange, current: number): number {
  if (range.kind === 'all') return 1
  if (range.kind === 'current') return current + 1
  return Math.max(Math.trunc(range.from) || 1, 1)
}
