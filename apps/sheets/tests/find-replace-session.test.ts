import { describe, expect, it } from 'vitest'
import type { IFindMatch, IFindReplaceService } from '@univerjs/find-replace'
import type { Workbook } from '@univerjs/core'

import {
  buildFindAllRows,
  FIND_HISTORY_KEY,
  getSessionMatches,
  hasUncommittedQuery,
  hideReplace,
  HISTORY_LIMIT,
  loadHistory,
  pushHistory,
} from '../src/renderer/find-replace-session'

function memoryStorage(): Pick<Storage, 'getItem' | 'setItem'> & { data: Map<string, string> } {
  const data = new Map<string, string>()
  return {
    data,
    getItem: (key) => data.get(key) ?? null,
    setItem: (key, value) => void data.set(key, value),
  }
}

describe('history', () => {
  it('round-trips committed queries, newest first, deduplicated', () => {
    const storage = memoryStorage()
    pushHistory(storage, FIND_HISTORY_KEY, 'alpha')
    pushHistory(storage, FIND_HISTORY_KEY, 'beta')
    pushHistory(storage, FIND_HISTORY_KEY, 'alpha')
    expect(loadHistory(storage, FIND_HISTORY_KEY)).toEqual(['alpha', 'beta'])
  })

  it('ignores blanks, caps the list and survives malformed storage', () => {
    const storage = memoryStorage()
    expect(pushHistory(storage, FIND_HISTORY_KEY, '   ')).toEqual([])
    for (let index = 0; index < HISTORY_LIMIT + 5; index += 1) {
      pushHistory(storage, FIND_HISTORY_KEY, `q${index}`)
    }
    expect(loadHistory(storage, FIND_HISTORY_KEY)).toHaveLength(HISTORY_LIMIT)
    storage.data.set(FIND_HISTORY_KEY, '{not json')
    expect(loadHistory(storage, FIND_HISTORY_KEY)).toEqual([])
    expect(loadHistory(undefined, FIND_HISTORY_KEY)).toEqual([])
  })
})

function fakeWorkbook(): Workbook {
  const cells: Record<string, { v?: unknown; f?: string }> = {
    's1:0:0': { v: 'total' },
    's1:2:1': { v: 42 },
    's1:3:1': { f: '=SUM(A1:A3)', v: null },
  }
  return {
    getUnitId: () => 'unit-1',
    getSheetBySheetId: (sheetId: string) =>
      sheetId === 's1' || sheetId === 's2'
        ? {
            getName: () => (sheetId === 's1' ? 'Data' : 'Other'),
            getCellRaw: (row: number, column: number) => cells[`${sheetId}:${row}:${column}`],
          }
        : null,
  } as unknown as Workbook
}

function cellMatch(
  unitId: string,
  subUnitId: string,
  row: number,
  column: number,
  matchedText?: string | null,
): IFindMatch {
  return {
    provider: 'test',
    unitId,
    range: {
      subUnitId,
      range: { startRow: row, startColumn: column, endRow: row, endColumn: column },
    },
    ...(matchedText !== undefined ? { matchedText } : {}),
  } as unknown as IFindMatch
}

describe('buildFindAllRows', () => {
  it('shapes matches into sheet/address/content rows', () => {
    const { rows, total } = buildFindAllRows(
      [
        cellMatch('unit-1', 's1', 0, 0),
        cellMatch('unit-1', 's1', 2, 1),
        cellMatch('unit-1', 's1', 3, 1),
        cellMatch('unit-1', 's2', 5, 3, 'streamed text'),
      ],
      fakeWorkbook(),
      500,
    )
    expect(total).toBe(4)
    expect(rows).toEqual([
      {
        sheetId: 's1',
        sheetName: 'Data',
        address: 'A1',
        content: 'total',
        bounds: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 },
      },
      {
        sheetId: 's1',
        sheetName: 'Data',
        address: 'B3',
        content: '42',
        bounds: { startRow: 2, startColumn: 1, endRow: 2, endColumn: 1 },
      },
      {
        sheetId: 's1',
        sheetName: 'Data',
        address: 'B4',
        content: '=SUM(A1:A3)',
        bounds: { startRow: 3, startColumn: 1, endRow: 3, endColumn: 1 },
      },
      {
        sheetId: 's2',
        sheetName: 'Other',
        address: 'D6',
        content: 'streamed text',
        bounds: { startRow: 5, startColumn: 3, endRow: 5, endColumn: 3 },
      },
    ])
  })

  it('skips other units and malformed matches but counts past the display limit', () => {
    const matches = [
      { provider: 'test', unitId: 'unit-2' } as IFindMatch,
      { provider: 'doc', unitId: 'unit-1', range: { startRow: 1 } } as unknown as IFindMatch,
      cellMatch('unit-1', 's1', 0, 0),
      cellMatch('unit-1', 's1', 2, 1),
    ]
    const { rows, total } = buildFindAllRows(matches, fakeWorkbook(), 1)
    expect(total).toBe(2)
    expect(rows).toHaveLength(1)
    expect(rows[0]?.address).toBe('A1')
  })
})

describe('service internals glue', () => {
  it('reads session matches and tolerates a missing model', () => {
    const matches = [cellMatch('unit-1', 's1', 0, 0)]
    const service = { _model: { _matches: matches } } as unknown as IFindReplaceService
    expect(getSessionMatches(service)).toBe(matches)
    expect(getSessionMatches({} as IFindReplaceService)).toEqual([])
    expect(getSessionMatches({ _model: {} } as unknown as IFindReplaceService)).toEqual([])
  })

  it('hideReplace flips the context flag, then the state, and survives their absence', () => {
    const calls: string[] = []
    const service = {
      _state: {
        changeState: (update: Record<string, unknown>) =>
          void calls.push(`state:${JSON.stringify(update)}`),
      },
      _toggleRevealReplace: (revealed: boolean) => void calls.push(`toggle:${revealed}`),
    } as unknown as IFindReplaceService
    hideReplace(service)
    expect(calls).toEqual(['toggle:false', 'state:{"replaceRevealed":false}'])
    expect(() => hideReplace({} as IFindReplaceService)).not.toThrow()
  })

  it('hasUncommittedQuery compares the typed query with the searched one', () => {
    expect(hasUncommittedQuery({ findString: 'apple', inputtingFindString: 'apple' })).toBe(false)
    expect(hasUncommittedQuery({ findString: 'apple', inputtingFindString: 'apples' })).toBe(true)
    expect(hasUncommittedQuery({ findString: '', inputtingFindString: '' })).toBe(false)
    expect(hasUncommittedQuery({ findString: 'apple', inputtingFindString: '' })).toBe(true)
  })
})
