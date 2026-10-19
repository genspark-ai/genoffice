import { describe, expect, it } from 'vitest'

import type { LazyWorkbookState } from '../src/renderer/univer-state'
import {
  largeFilterHiddenRows,
  largeFilterOwnsSpan,
  largeFilterSheet,
  resetLargeFilter,
} from '../src/renderer/large-filter-state'
import {
  buildFilterPredicate,
  buildFilterValueTree,
  hiddenRowsForColumn,
  largeFilterAction,
  projectionReader,
  projectionText,
  virtualWorksheet,
} from '../src/renderer/large-filter'

function stateWithJournal(
  cells: Record<string, { value: string | number | null; hasValue?: boolean }> = {},
): LazyWorkbookState {
  const sheetCells = new Map(
    Object.entries(cells).map(([key, entry]) => [
      key,
      { hasValue: entry.hasValue ?? true, value: entry.value, row: 0, column: 0 },
    ]),
  )
  return {
    editJournal: {
      cells: new Map([['s1', sheetCells]]),
      bulkConstantFills: new Map(),
      structuralOps: new Map(),
    },
  } as unknown as LazyWorkbookState
}

describe('projectionText', () => {
  it('matches the text the by-values filter compares', () => {
    expect(projectionText(null)).toBe('')
    expect(projectionText('a\r\nb')).toBe('ab')
    expect(projectionText(12.5)).toBe('12.5')
    expect(projectionText(true)).toBe('TRUE')
  })
})

describe('buildFilterPredicate', () => {
  it('filters by values with blanks opt-in', () => {
    const keep = buildFilterPredicate({ colId: 0, filters: { filters: ['BP000010', '7'] } })!
    expect(keep('BP000010')).toBe(true)
    expect(keep(7)).toBe(true)
    expect(keep('')).toBe(false)
    expect(keep(null)).toBe(false)
    const withBlank = buildFilterPredicate({ colId: 0, filters: { blank: true, filters: ['x'] } })!
    expect(withBlank(null)).toBe(true)
  })

  it('evaluates numeric conditions on raw numbers and wildcards on text', () => {
    const gt = buildFilterPredicate({
      colId: 0,
      customFilters: { customFilters: [{ val: 100, operator: 'greaterThan' as never }] },
    })!
    expect(gt(101)).toBe(true)
    expect(gt(100)).toBe(false)
    expect(gt('abc')).toBe(false)
    const starts = buildFilterPredicate({
      colId: 0,
      customFilters: { customFilters: [{ val: 'BP0*' }] },
    })!
    expect(starts('BP000001')).toBe(true)
    expect(starts('\u4e3b\u4f53')).toBe(false)
    const between = buildFilterPredicate({
      colId: 0,
      customFilters: {
        and: 1 as never,
        customFilters: [
          { val: 10, operator: 'greaterThanOrEqual' as never },
          { val: 20, operator: 'lessThan' as never },
        ],
      },
    })!
    expect(between(10)).toBe(true)
    expect(between(20)).toBe(false)
    const notX = buildFilterPredicate({
      colId: 0,
      customFilters: { customFilters: [{ val: 'x', operator: 'notEqual' as never }] },
    })!
    expect(notX('y')).toBe(true)
    expect(notX('x')).toBe(false)
  })

  it('declines color filters', () => {
    expect(
      buildFilterPredicate({ colId: 0, colorFilters: { cellFillColors: ['#fff'] } }),
    ).toBeNull()
  })
})

describe('hiddenRowsForColumn', () => {
  it('hides rows failing the predicate, with session edits overriding the file', () => {
    const projection = { startRow: 5, values: ['a', 'b', 'a', null], styleIndex: [] }
    const state = stateWithJournal({ '6:2': { value: 'a' }, '8:2': { value: 'zzz' } })
    const read = projectionReader(state, 's1', 2, projection)
    const keepA = buildFilterPredicate({ colId: 2, filters: { filters: ['a'] } })!
    expect([...hiddenRowsForColumn(read, keepA, 5, 8)]).toEqual([8])
  })
})

describe('virtualWorksheet', () => {
  it('serves projection values as cells with the date pattern style', () => {
    const projection = {
      startRow: 1,
      values: [45000, 'text', true, null],
      styleIndex: [0, undefined, undefined, undefined],
    }
    const read = projectionReader(stateWithJournal(), 's1', 0, projection)
    const sheet = virtualWorksheet(read, (row) =>
      row === 1 ? { n: { pattern: 'yyyy-mm-dd' } } : undefined,
    )
    const cells = [
      ...sheet.iterateByColumn({ startRow: 1, endRow: 4, startColumn: 0, endColumn: 0 }),
    ]
    expect(cells.map((c) => c.value)).toEqual([
      { v: 45000, s: { n: { pattern: 'yyyy-mm-dd' } } },
      { v: 'text' },
      { v: 1, t: 3 },
      null,
    ])
    expect(sheet.getCellRaw(1, 0)).toEqual({ v: 45000, s: { n: { pattern: 'yyyy-mm-dd' } } })
  })
})

describe('largeFilterAction', () => {
  const worksheet = {} as never
  it('takes over criteria commands and leaves panel/range commands to Univer', () => {
    expect(largeFilterAction('sheet.operation.open-filter-panel', {}, worksheet)).toBeNull()
    expect(largeFilterAction('sheet.command.smart-toggle-filter', {}, worksheet)).toBeNull()
    expect(largeFilterAction('sheet.command.remove-sheet-filter', {}, worksheet)).toBeNull()
    expect(
      largeFilterAction('sheet.command.set-filter-criteria', { col: 2, criteria: null }, worksheet),
    ).toBeTypeOf('function')
    expect(largeFilterAction('sheet.command.set-filter-criteria', {}, worksheet)).toBeNull()
    expect(largeFilterAction('sheet.command.clear-filter-criteria', {}, worksheet)).toBeTypeOf(
      'function',
    )
    expect(largeFilterAction('sheet.command.re-calc-filter', {}, worksheet)).toBeTypeOf('function')
  })
})

describe('buildFilterValueTree', () => {
  const t = (key: string) =>
    key === 'sheets-filter-ui.panel.empty' ? '(empty)' : key.split('.').pop()!
  const styles = { get: (s: unknown) => (typeof s === 'string' ? undefined : (s as never)) }
  const range = { startRow: 1, endRow: 6, startColumn: 0, endColumn: 0 }

  it('counts distinct texts, blanks and date groups like the stock panel', () => {
    const projection = {
      startRow: 1,
      values: ['b', 'a', null, 'a', 45000, 45001],
      styleIndex: [undefined, undefined, undefined, undefined, 0, 0],
    }
    const read = projectionReader(stateWithJournal(), 's1', 0, projection)
    const source = virtualWorksheet(read, (row) =>
      row >= 5 ? { n: { pattern: 'yyyy-mm-dd' } } : undefined,
    )
    const tree = buildFilterValueTree({
      source,
      iterateRange: range,
      styles,
      filteredOutRowsByOtherColumns: new Set(),
      hasFilters: false,
      filterBy: undefined,
      filteredOutCount: 0,
      alreadyChecked: new Set(),
      blankChecked: false,
      t,
    })
    const flat = tree.filterTreeItems.map((i) => `${i.key}:${i.count}:${i.checked}`)
    expect(flat).toEqual(['2023:2:false', 'empty:1:true', 'a:2:true', 'b:1:true'])
    const year = tree.filterTreeItems[0]!
    expect(year.children?.map((m) => `${m.key}:${m.count}`)).toEqual(['2023-3:2'])
    expect(year.children?.[0]?.children?.map((d) => d.key)).toEqual(['2023-3-15', '2023-3-16'])
    expect(tree.filterTreeMapCache.get('2023-3-16')).toEqual(['2023', '2023-3', '2023-3-16'])
  })

  it('skips rows hidden by other columns and keeps the checked set', () => {
    const projection = { startRow: 1, values: ['x', 'y', 'x', 'z', null, null], styleIndex: [] }
    const source = virtualWorksheet(
      projectionReader(stateWithJournal(), 's1', 0, projection),
      () => undefined,
    )
    const tree = buildFilterValueTree({
      source,
      iterateRange: range,
      styles,
      filteredOutRowsByOtherColumns: new Set([2, 6]),
      hasFilters: true,
      filterBy: 0,
      filteredOutCount: 3,
      alreadyChecked: new Set(['x']),
      blankChecked: false,
      t,
    })
    expect(tree.filterTreeItems.map((i) => `${i.key}:${i.count}:${i.checked}`)).toEqual([
      'empty:1:false',
      'x:2:true',
      'z:1:false',
    ])
  })
})

describe('resetLargeFilter', () => {
  it('releases the span and the hidden set once the AutoFilter is removed', () => {
    const state = stateWithJournal()
    const entry = largeFilterSheet(state, 's1')
    entry.active = true
    entry.hiddenByColumn.set(0, new Set([3, 4]))
    expect(largeFilterOwnsSpan(state, 's1')).toBe(true)
    expect([...largeFilterHiddenRows(state, 's1')!]).toEqual([3, 4])
    resetLargeFilter(state, 's1')
    expect(largeFilterOwnsSpan(state, 's1')).toBe(false)
    expect(largeFilterHiddenRows(state, 's1')).toBeUndefined()
    expect(entry.hiddenByColumn.size).toBe(0)
  })
})

describe('resetLargeFilter generation', () => {
  it('bumps so an apply that was in flight discards its result', () => {
    const state = stateWithJournal()
    const entry = largeFilterSheet(state, 's1')
    const before = entry.generation
    resetLargeFilter(state, 's1')
    expect(entry.generation).toBe(before + 1)
  })
})
