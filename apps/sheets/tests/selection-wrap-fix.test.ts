import { Direction } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import {
  MOVE_SELECTION_COMMAND,
  installSelectionWrapGuard,
  moveWouldWrap,
  type SelectionRange,
  type WorksheetGrid,
} from '../src/renderer/selection-wrap-fix'
import type { UniverRuntime } from '../src/renderer/univer-state'

function grid(
  rows: number,
  cols: number,
  hiddenRows: number[] = [],
  hiddenCols: number[] = [],
): WorksheetGrid {
  const hr = new Set(hiddenRows)
  const hc = new Set(hiddenCols)
  return {
    getRowCount: () => rows,
    getColumnCount: () => cols,
    getRowVisible: (row) => !hr.has(row),
    getColVisible: (col) => !hc.has(col),
  }
}

function cell(row: number, col: number): SelectionRange {
  return { startRow: row, endRow: row, startColumn: col, endColumn: col }
}

describe('moveWouldWrap', () => {
  it('wraps at each sheet edge', () => {
    const sheet = grid(1000, 26)
    expect(moveWouldWrap(Direction.DOWN, cell(999, 25), sheet)).toBe(true)
    expect(moveWouldWrap(Direction.UP, cell(0, 0), sheet)).toBe(true)
    expect(moveWouldWrap(Direction.RIGHT, cell(0, 25), sheet)).toBe(true)
    expect(moveWouldWrap(Direction.LEFT, cell(0, 0), sheet)).toBe(true)
  })

  it('does not trigger mid-sheet', () => {
    const sheet = grid(1000, 26)
    expect(moveWouldWrap(Direction.DOWN, cell(500, 5), sheet)).toBe(false)
    expect(moveWouldWrap(Direction.UP, cell(500, 5), sheet)).toBe(false)
    expect(moveWouldWrap(Direction.LEFT, cell(500, 5), sheet)).toBe(false)
    expect(moveWouldWrap(Direction.RIGHT, cell(500, 5), sheet)).toBe(false)
  })

  it('treats trailing hidden rows/columns as the edge', () => {
    const sheet = grid(10, 10, [8, 9], [9])
    expect(moveWouldWrap(Direction.DOWN, cell(7, 0), sheet)).toBe(true)
    expect(moveWouldWrap(Direction.DOWN, cell(6, 0), sheet)).toBe(false)
    expect(moveWouldWrap(Direction.RIGHT, cell(0, 8), sheet)).toBe(true)
    expect(moveWouldWrap(Direction.UP, cell(1, 0), grid(10, 10, [0])), 'hidden first row').toBe(
      true,
    )
  })

  it('uses the selection edge nearest the move direction', () => {
    const sheet = grid(10, 10)
    const range: SelectionRange = { startRow: 5, endRow: 9, startColumn: 0, endColumn: 2 }
    expect(moveWouldWrap(Direction.DOWN, range, sheet)).toBe(true)
    expect(moveWouldWrap(Direction.UP, range, sheet)).toBe(false)
  })
})

interface FakeCommandEvent {
  id: string
  params?: unknown
  cancel?: boolean
}

function fakeRuntime(range: SelectionRange | null, sheet: WorksheetGrid) {
  const handlers: ((event: FakeCommandEvent) => void)[] = []
  const univerAPI = {
    Event: { BeforeCommandExecute: 'BeforeCommandExecute' },
    addEvent: (_name: string, handler: (event: FakeCommandEvent) => void) => {
      handlers.push(handler)
      return { dispose() {} }
    },
    getActiveWorkbook: () => ({
      getActiveRange: () => (range ? { getRange: () => range } : null),
      getActiveSheet: () => ({ getSheet: () => sheet }),
    }),
  }
  return {
    runtime: { univerAPI } as unknown as UniverRuntime,
    emit(event: FakeCommandEvent): FakeCommandEvent {
      handlers.forEach((handler) => handler(event))
      return event
    },
  }
}

describe('installSelectionWrapGuard', () => {
  it('cancels an arrow move at the edge', () => {
    const { runtime, emit } = fakeRuntime(cell(999, 0), grid(1000, 26))
    installSelectionWrapGuard(runtime)
    const event = emit({ id: MOVE_SELECTION_COMMAND, params: { direction: Direction.DOWN } })
    expect(event.cancel).toBe(true)
  })

  it('also guards jumpOver (Ctrl+arrow) at the edge', () => {
    const { runtime, emit } = fakeRuntime(cell(0, 25), grid(1000, 26))
    installSelectionWrapGuard(runtime)
    const event = emit({
      id: MOVE_SELECTION_COMMAND,
      params: { direction: Direction.RIGHT, jumpOver: 1 },
    })
    expect(event.cancel).toBe(true)
  })

  it('lets mid-sheet moves and other commands through', () => {
    const { runtime, emit } = fakeRuntime(cell(10, 10), grid(1000, 26))
    installSelectionWrapGuard(runtime)
    expect(
      emit({ id: MOVE_SELECTION_COMMAND, params: { direction: Direction.DOWN } }).cancel,
    ).toBeUndefined()
    expect(emit({ id: 'sheet.command.set-range-values', params: {} }).cancel).toBeUndefined()
  })

  it('skips when there is no active range', () => {
    const { runtime, emit } = fakeRuntime(null, grid(10, 10))
    installSelectionWrapGuard(runtime)
    const event = emit({ id: MOVE_SELECTION_COMMAND, params: { direction: Direction.UP } })
    expect(event.cancel).toBeUndefined()
  })
})
