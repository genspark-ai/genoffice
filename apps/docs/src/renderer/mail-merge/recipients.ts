import type { MergeRecord } from '@genoffice/docx-engine'

/**
 * The view of the recipient list: records with their source indexes, narrowed
 * by the filter conditions and sorted by the column the user picked — Word's
 * Filter and Sort in the mail-merge recipients dialog.
 *
 * The source index travels with each row, so the merge still takes records out
 * of the list after filtering and sorting.
 */
export interface RecipientRow {
  /** the record's index in the data source */
  index: number
  record: MergeRecord
}

/** the comparisons Word offers when filtering the recipient list */
export type FilterOperator =
  'eq' | 'ne' | 'gt' | 'lt' | 'gte' | 'lte' | 'contains' | 'notContains' | 'blank' | 'notBlank'

/** how a condition joins the one above it */
export type FilterJoin = 'and' | 'or'

/**
 * One filter condition. A condition with an empty field is inactive: the
 * filter panel keeps a blank row ready, and it does not narrow anything until
 * the user picks a field.
 */
export interface FilterCondition {
  field: string
  operator: FilterOperator
  value: string
  join: FilterJoin
}

export type SortDirection = 'asc' | 'desc'

/** one sort level: a column and its direction; '' keeps the source order */
export interface SortKey {
  field: string
  direction: SortDirection
}

const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' })

/** the number a value reads as, or null when it is not a plain number */
function asNumber(value: string): number | null {
  const trimmed = value.trim()
  if (trimmed === '') return null
  const n = Number(trimmed)
  return Number.isNaN(n) ? null : n
}

/**
 * Compares two cell values the way Word does: numbers compare as numbers, so
 * "9" is less than "10"; everything else compares as text.
 */
function compareValues(a: string, b: string): number {
  const na = asNumber(a)
  const nb = asNumber(b)
  if (na !== null && nb !== null) return na - nb
  return collator.compare(a, b)
}

function matches(record: MergeRecord, condition: FilterCondition): boolean {
  const cell = record[condition.field] ?? ''
  const wanted = condition.value
  switch (condition.operator) {
    case 'eq':
      return cell === wanted
    case 'ne':
      return cell !== wanted
    case 'gt':
      return compareValues(cell, wanted) > 0
    case 'lt':
      return compareValues(cell, wanted) < 0
    case 'gte':
      return compareValues(cell, wanted) >= 0
    case 'lte':
      return compareValues(cell, wanted) <= 0
    case 'contains':
      return cell.toLowerCase().includes(wanted.toLowerCase())
    case 'notContains':
      return !cell.toLowerCase().includes(wanted.toLowerCase())
    case 'blank':
      return cell.trim() === ''
    case 'notBlank':
      return cell.trim() !== ''
  }
}

export function viewRecipients(
  records: readonly MergeRecord[],
  conditions: readonly FilterCondition[],
  sorts: readonly SortKey[],
): RecipientRow[] {
  let rows = records.map((record, index) => ({ index, record }))
  const active = conditions.filter((c) => c.field !== '')
  if (active.length > 0) {
    // conditions fold left to right: the first one decides, then each join
    // picks how the one below it joins in
    rows = rows.filter(({ record }) => {
      let keep = matches(record, active[0])
      for (let i = 1; i < active.length; i++) {
        const hit = matches(record, active[i])
        keep = active[i].join === 'and' ? keep && hit : keep || hit
      }
      return keep
    })
  }
  // sort levels apply in order: the primary field first, then the two fallbacks
  const levels = sorts.filter((s) => s.field !== '')
  if (levels.length > 0) {
    rows = [...rows].sort((a, b) => {
      for (const { field, direction } of levels) {
        const diff = compareValues(a.record[field] ?? '', b.record[field] ?? '')
        if (diff !== 0) return direction === 'asc' ? diff : -diff
      }
      return 0
    })
  }
  return rows
}
