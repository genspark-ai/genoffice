import { describe, expect, it } from 'vitest'
import {
  activeRecords,
  compareRecipients,
  findRecipient,
  matchesConditions,
  matchesFilter,
  patchQueryRow,
  recipientConditions,
  recipientFilter,
  type RecipientCondition,
  type RecipientQuery,
} from '../src/renderer/mail-merge/recipients'

const RECORDS = [
  { Name: 'Ada', City: 'London' },
  { Name: 'Grace', City: 'New York' },
  { Name: 'Linus', City: '' },
  { Name: 'Alan', City: 'london' },
]

describe('mail merge recipients', () => {
  it('a filter compares without regard to case', () => {
    const city = (op: Parameters<typeof matchesFilter>[1]['op'], value = '') =>
      RECORDS.map((record) => matchesFilter(record, { field: 'City', op, value }))
    expect(city('equals', 'LONDON')).toEqual([true, false, false, true])
    expect(city('notEquals', 'London')).toEqual([false, true, true, false])
    expect(city('contains', 'york')).toEqual([false, true, false, false])
    expect(city('notContains', 'on')).toEqual([false, true, true, false])
    expect(city('blank')).toEqual([false, false, true, false])
    expect(city('notBlank')).toEqual([true, true, false, true])
  })

  it('a form left without a value to compare to is no condition', () => {
    expect(recipientFilter('City', 'equals', '')).toBeNull()
    expect(recipientFilter('City', 'notContains', '  ')).toBeNull()
    expect(recipientFilter('', 'blank', '')).toBeNull()
    expect(recipientFilter('City', 'contains', 'on')).toEqual({
      field: 'City',
      op: 'contains',
      value: 'on',
    })
    // blank and not-blank take no value, and drop one typed before the switch
    expect(recipientFilter('City', 'blank', 'x')).toEqual({ field: 'City', op: 'blank', value: '' })
  })

  it('the merge takes the ticked rows that pass the filter', () => {
    expect(activeRecords(RECORDS, new Set())).toEqual(RECORDS)
    expect(activeRecords(RECORDS, new Set([0])).map((r) => r.Name)).toEqual([
      'Grace',
      'Linus',
      'Alan',
    ])
    const london = { join: 'and', field: 'City', op: 'equals', value: 'london' } as const
    expect(
      activeRecords(RECORDS, new Set([0]), { conditions: [london], sort: [] }).map((r) => r.Name),
    ).toEqual(['Alan'])
  })

  it('And binds before Or', () => {
    const is = (
      join: 'and' | 'or',
      field: string,
      value: string,
      op: RecipientCondition['op'] = 'equals',
    ): RecipientCondition => ({ join, field, op, value })
    const names = (...conditions: RecipientCondition[]) =>
      RECORDS.filter((record) => matchesConditions(record, conditions)).map((r) => r.Name)
    expect(names()).toEqual(['Ada', 'Grace', 'Linus', 'Alan'])
    expect(names(is('and', 'City', 'london'), is('and', 'Name', 'a', 'contains'))).toEqual([
      'Ada',
      'Alan',
    ])
    expect(names(is('and', 'City', 'london'), is('and', 'Name', 'Ada'))).toEqual(['Ada'])
    expect(names(is('and', 'City', 'new york'), is('or', 'Name', 'Ada'))).toEqual(['Ada', 'Grace'])
    // Grace, or (London and Alan): not (Grace or London) and Alan
    expect(
      names(is('and', 'Name', 'Grace'), is('or', 'City', 'london'), is('and', 'Name', 'Alan')),
    ).toEqual(['Grace', 'Alan'])
    // (London and Ada) or (blank city): the And group comes first as well
    expect(
      names(is('and', 'City', 'london'), is('and', 'Name', 'Ada'), is('or', 'City', '', 'blank')),
    ).toEqual(['Ada', 'Linus'])
    // the first condition's join is not read
    expect(names(is('or', 'Name', 'Ada'), is('and', 'City', 'london'))).toEqual(['Ada'])
  })

  it('a row left without a value is dropped from the conditions, the others stay', () => {
    const row = (
      join: 'and' | 'or',
      field: string,
      op: RecipientCondition['op'],
      value: string,
    ) => ({
      join,
      field,
      op,
      value,
    })
    expect(
      recipientConditions([row('and', 'City', 'equals', ''), row('and', '', 'equals', 'x')]),
    ).toEqual([])
    const conditions = recipientConditions([
      row('and', 'City', 'equals', 'London'),
      row('or', 'Name', 'contains', '  '),
      row('or', 'City', 'blank', 'typed before the switch'),
    ])
    expect(conditions).toEqual([
      { join: 'and', field: 'City', op: 'equals', value: 'London' },
      { join: 'or', field: 'City', op: 'blank', value: '' },
    ])
    // the dropped row filters nothing out
    expect(activeRecords(RECORDS, new Set(), { conditions, sort: [] }).map((r) => r.Name)).toEqual([
      'Ada',
      'Linus',
      'Alan',
    ])
  })

  it('Sort Records orders the recipients the filter leaves, ties keeping the list order', () => {
    const people = [
      { Name: 'Grace', City: 'Paris', Age: '41' },
      { Name: 'ada', City: 'London', Age: '36' },
      { Name: 'Alan', City: 'London', Age: '9' },
      { Name: 'Linus', City: 'Paris', Age: '30' },
      { Name: 'Bo', City: 'London', Age: '36' },
    ]
    const sorted = (sort: RecipientQuery['sort'], excluded = new Set<number>()) =>
      activeRecords(people, excluded, { conditions: [], sort }).map((r) => r.Name)
    expect(sorted([{ field: 'Name', descending: false }])).toEqual([
      'ada',
      'Alan',
      'Bo',
      'Grace',
      'Linus',
    ])
    // numbers by their value, the largest first; the two 36s stay as listed
    expect(sorted([{ field: 'Age', descending: true }])).toEqual([
      'Grace',
      'ada',
      'Bo',
      'Linus',
      'Alan',
    ])
    expect(
      sorted(
        [
          { field: 'City', descending: true },
          { field: 'Age', descending: false },
        ],
        new Set([0]),
      ),
    ).toEqual(['Linus', 'Alan', 'ada', 'Bo'])
    expect(compareRecipients([], people[0], people[1])).toBe(0)
    // the list itself is left as it was read
    expect(people[0].Name).toBe('Grace')
  })

  it('Find Recipient moves to the next record holding the text and wraps', () => {
    expect(findRecipient(RECORDS, 0, 'london', null)).toBe(3)
    expect(findRecipient(RECORDS, 3, 'london', null)).toBe(0)
    // the only match is the record it starts on
    expect(findRecipient(RECORDS, 1, 'grace', 'Name')).toBe(1)
    expect(findRecipient(RECORDS, 0, 'a', 'City')).toBe(-1)
    expect(findRecipient(RECORDS, 0, '  ', null)).toBe(-1)
    expect(findRecipient([], 0, 'x', null)).toBe(-1)
  })

  it('clearing the field of a query row empties the rows it locks', () => {
    const blank = { join: 'and' as const, field: '', op: 'equals' as const, value: '' }
    const rows = [
      { ...blank, field: 'City', value: 'London' },
      { ...blank, join: 'or' as const, field: 'Name', value: 'Ada' },
      { ...blank, field: 'City', op: 'blank' as const },
    ]
    const cleared = patchQueryRow(rows, 1, { field: '' }, blank)
    expect(cleared).toEqual([rows[0], { ...rows[1], field: '' }, blank])
    expect(recipientConditions(cleared)).toHaveLength(1)
    // any other edit leaves the rows under it alone
    expect(patchQueryRow(rows, 0, { field: 'Name' }, blank).slice(1)).toEqual(rows.slice(1))
    expect(patchQueryRow(rows, 0, { value: '' }, blank).slice(1)).toEqual(rows.slice(1))
  })
})
