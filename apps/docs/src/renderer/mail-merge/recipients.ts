/** the recipients a merge takes from its list: the rows left ticked that pass the filter, in the order asked for */

export type RecipientOp = 'equals' | 'notEquals' | 'contains' | 'notContains' | 'blank' | 'notBlank'

export const RECIPIENT_OPS: readonly RecipientOp[] = [
  'equals',
  'notEquals',
  'contains',
  'notContains',
  'blank',
  'notBlank',
]

export interface RecipientFilter {
  field: string
  op: RecipientOp
  value: string
}

/** how a condition joins the ones above it; the first condition's is not read */
export type RecipientJoin = 'and' | 'or'

export interface RecipientCondition extends RecipientFilter {
  join: RecipientJoin
}

export interface RecipientSortKey {
  field: string
  descending: boolean
}

/** Word's Query Options: Filter Records and Sort Records */
export interface RecipientQuery {
  conditions: readonly RecipientCondition[]
  sort: readonly RecipientSortKey[]
}

export const MAX_RECIPIENT_CONDITIONS = 6
export const MAX_RECIPIENT_SORT_KEYS = 3
export const EMPTY_RECIPIENT_QUERY: RecipientQuery = { conditions: [], sort: [] }

type MergeRecord = Record<string, string>

/** a comparison against nothing is no condition: blank and not-blank are the ones that take no value */
export const opTakesValue = (op: RecipientOp) => op !== 'blank' && op !== 'notBlank'

/** the condition a filled-in form stands for; null when it would compare against an empty value */
export function recipientFilter(
  field: string,
  op: RecipientOp,
  value: string,
): RecipientFilter | null {
  if (!field) return null
  if (!opTakesValue(op)) return { field, op, value: '' }
  return value.trim() ? { field, op, value } : null
}

const fold = (text: string) => text.trim().toLocaleLowerCase()

/** Word compares without regard to case */
export function matchesFilter(record: MergeRecord, filter: RecipientFilter): boolean {
  const cell = fold(record[filter.field] ?? '')
  const value = fold(filter.value)
  switch (filter.op) {
    case 'equals':
      return cell === value
    case 'notEquals':
      return cell !== value
    case 'contains':
      return cell.includes(value)
    case 'notContains':
      return !cell.includes(value)
    case 'blank':
      return cell === ''
    case 'notBlank':
      return cell !== ''
  }
}

/** Query Options row edit: a row left without a field locks the rows under it, so they empty with it */
export function patchQueryRow<T extends { field: string }>(
  rows: readonly T[],
  index: number,
  patch: Partial<T>,
  blank: T,
): T[] {
  return rows.map((row, i) =>
    i === index ? { ...row, ...patch } : patch.field === '' && i > index ? { ...blank } : row,
  )
}

/** the rows of a filled-in form that are conditions: one comparing against nothing is dropped */
export function recipientConditions(
  rows: ReadonlyArray<{ join: RecipientJoin; field: string; op: RecipientOp; value: string }>,
): RecipientCondition[] {
  return rows.flatMap((row) => {
    const filter = recipientFilter(row.field, row.op, row.value)
    return filter ? [{ ...filter, join: row.join }] : []
  })
}

/**
 * And binds before Or, as in Word: `a Or b And c` takes a record that meets a,
 * or both b and c. No condition at all takes every record.
 */
export function matchesConditions(
  record: MergeRecord,
  conditions: readonly RecipientCondition[],
): boolean {
  if (conditions.length === 0) return true
  let all = true
  for (const [index, condition] of conditions.entries()) {
    if (index > 0 && condition.join === 'or') {
      if (all) return true
      all = true
    }
    all = all && matchesFilter(record, condition)
  }
  return all
}

const collator = new Intl.Collator(undefined, { sensitivity: 'base', numeric: true })

/** the first key that tells two records apart decides; a tie on every key keeps the list's order */
export function compareRecipients(
  keys: readonly RecipientSortKey[],
  a: MergeRecord,
  b: MergeRecord,
): number {
  for (const key of keys) {
    const result = collator.compare((a[key.field] ?? '').trim(), (b[key.field] ?? '').trim())
    if (result !== 0) return key.descending ? -result : result
  }
  return 0
}

/** `excluded` holds positions in `records`, the rows the user unticked */
export function activeRecords(
  records: readonly MergeRecord[],
  excluded: ReadonlySet<number>,
  query: RecipientQuery = EMPTY_RECIPIENT_QUERY,
): MergeRecord[] {
  const kept = records.filter(
    (record, index) => !excluded.has(index) && matchesConditions(record, query.conditions),
  )
  return query.sort.length > 0 ? kept.sort((a, b) => compareRecipients(query.sort, a, b)) : kept
}

/**
 * The next record after `from` holding `text`, in one field or in any, wrapping
 * past the end; -1 when no record holds it.
 */
export function findRecipient(
  records: readonly MergeRecord[],
  from: number,
  text: string,
  field: string | null,
): number {
  const needle = fold(text)
  if (!needle) return -1
  for (let step = 1; step <= records.length; step++) {
    const index = (from + step) % records.length
    const record = records[index]
    const cells = field ? [record[field] ?? ''] : Object.values(record)
    if (cells.some((cell) => fold(cell).includes(needle))) return index
  }
  return -1
}
