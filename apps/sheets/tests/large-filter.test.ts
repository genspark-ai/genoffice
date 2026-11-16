import { describe, expect, it } from 'vitest'
import type { LazyWorkbookState } from '../src/renderer/univer-state'
import { filteredViewportRows } from '../src/renderer/univer-sync'
import {
  decodeFilterHiddenRows,
  filterRequest,
  largeFilterAction,
} from '../src/renderer/large-filter'
import {
  largeFilterHiddenRows,
  largeFilterSheet,
  resetLargeFilter,
} from '../src/renderer/large-filter-state'
import { distinctBoundaries, visibleCellRanges } from '../src/renderer/sparse-sheet-render'
import { screenToFile } from '../src/renderer/view-transform'

describe('streamed AutoFilter', () => {
  it('sends affine row spans, omits the listed column from criteria and layers journal edits', () => {
    const sheetId = 's1'
    const ops = [
      { kind: 'remove-rows', index: 8, count: 3 },
      { kind: 'insert-rows', index: 4, count: 2 },
      { kind: 'move-rows', index: 12, count: 4, before: 6 },
      { kind: 'remove-cols', index: 0, count: 1 },
    ] as const
    const state = {
      file: { sessionId: 'session', sheets: [{ id: sheetId, rowCount: 21 }] },
      editJournal: {
        structuralOps: new Map([[sheetId, ops]]),
        bulkConstantFills: new Map([
          [
            sheetId,
            [
              { startRow: 2, endRow: 8, startColumn: 1, endColumn: 1, value: 'fill' },
              { startRow: 3, endRow: 4, startColumn: 1, endColumn: 2, value: 'later' },
            ],
          ],
        ]),
        cells: new Map([
          [
            sheetId,
            new Map([
              ['4:1', { row: 4, column: 1, value: '', hasValue: true }],
              ['5:1', { row: 5, column: 1, value: null, hasValue: false }],
            ]),
          ],
        ]),
      },
    } as unknown as LazyWorkbookState
    const range = { startRow: 2, endRow: 18, startColumn: 0, endColumn: 1 }
    const request = filterRequest(
      state,
      sheetId,
      range,
      new Map([
        [0, { colId: 0, filters: { filters: ['keep'] } }],
        [1, { colId: 1, filters: { blank: true } }],
      ]),
      { column: 1, query: 'abc', offset: 200, limit: 100 },
    )
    expect(request.criteria).toEqual([
      {
        fileColumn: 1,
        overrides: [],
        predicate: { kind: 'values', values: ['keep'], blank: false },
      },
    ])
    expect(request.page?.column).toEqual({
      fileColumn: 2,
      overrides: [
        { startRow: 3, endRow: 8, value: 'fill' },
        { startRow: 3, endRow: 4, value: 'later' },
        { startRow: 4, endRow: 4, value: '' },
      ],
    })
    const mapped = new Map<number, number>()
    for (const span of request.spans) {
      for (let r = span.startRow; r <= span.endRow; r++) {
        expect(mapped.has(r)).toBe(false)
        mapped.set(r, span.fileStartRow + r - span.startRow)
      }
    }
    for (let r = request.startRow; r <= request.endRow; r++) {
      expect(mapped.get(r) ?? null).toBe(screenToFile(ops, 'row', r))
    }
  })

  it('selects the physical rows visible across large gaps and skips manual hides', () => {
    const last = 968_576
    const bytes = new Uint8Array(Math.ceil(last / 8)).fill(255)
    const matches = [221_269, 284_356, 936_617, last]
    for (const row of matches) bytes[(row - 1) >> 3]! &= ~(1 << ((row - 1) % 8))
    const hidden = decodeFilterHiddenRows(Buffer.from(bytes).toString('base64'), 1, last)
    expect(filteredViewportRows(0, last + 1, 20, hidden, (r) => r !== matches[1])).toEqual([
      0,
      matches[0],
      matches[2],
      last,
    ])
    expect(filteredViewportRows(matches[2]!, last + 1, 1, hidden, () => true)).toEqual([matches[2]])
  })

  it('draws each gridline once across runs of collapsed rows', () => {
    const positions = new Array<number>(1_000_001).fill(20)
    positions.fill(40, 221_268)
    positions.fill(60, 936_616)
    positions[1_000_000] = 80
    expect([...distinctBoundaries(positions, 0, 1_000_000)]).toEqual([20, 40, 60, 80])
    expect([...distinctBoundaries(positions, 100, 950_000)]).toEqual([20, 40, 60])
    expect([...distinctBoundaries(positions, 222_000, 222_100)]).toEqual([40])
    expect(
      visibleCellRanges(positions, {
        startRow: 1,
        endRow: 1_000_000,
        startColumn: 0,
        endColumn: 10,
      }),
    ).toEqual(
      [221_268, 936_616, 1_000_000].map((row) => ({
        startRow: row,
        endRow: row,
        startColumn: 0,
        endColumn: 10,
      })),
    )
  })

  it('releases visibility and invalidates an in-flight apply on removal', () => {
    const state = {} as LazyWorkbookState
    const entry = largeFilterSheet(state, 's1')
    entry.visibility = { filteredOutRows: new Set([3, 4]) }
    const generation = entry.generation
    expect(largeFilterHiddenRows(state, 's1')).toBe(entry.visibility.filteredOutRows)
    resetLargeFilter(state, 's1')
    expect(largeFilterHiddenRows(state, 's1')).toBeUndefined()
    expect(entry.generation).toBe(generation + 1)
  })

  it('intercepts criteria commands while leaving panel/range commands to Univer', () => {
    const sheet = {} as never
    expect(largeFilterAction('sheet.operation.open-filter-panel', {}, sheet)).toBeNull()
    expect(largeFilterAction('sheet.command.smart-toggle-filter', {}, sheet)).toBeNull()
    expect(largeFilterAction('sheet.command.set-filter-criteria', { col: 2 }, sheet)).toBeTypeOf(
      'function',
    )
    expect(largeFilterAction('sheet.command.clear-filter-criteria', {}, sheet)).toBeTypeOf(
      'function',
    )
    expect(largeFilterAction('sheet.command.re-calc-filter', {}, sheet)).toBeTypeOf('function')
  })
})
