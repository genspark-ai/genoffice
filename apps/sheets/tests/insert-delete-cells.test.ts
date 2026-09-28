import { describe, expect, it } from 'vitest'
import { RANGE_TYPE } from '@univerjs/core'

import {
  DEFAULT_CELLS_CHOICE,
  DELETE_RANGE_LEFT_COMMAND,
  DELETE_RANGE_UP_COMMAND,
  INSERT_COLS_COMMAND,
  INSERT_RANGE_DOWN_COMMAND,
  INSERT_RANGE_RIGHT_COMMAND,
  INSERT_ROWS_COMMAND,
  REMOVE_COLS_COMMAND,
  REMOVE_ROWS_COMMAND,
  cellsCommandId,
  cellsCommandParams,
  decideCellsAction,
  type SelectionShape,
} from '../src/renderer/insert-delete-cells'
import { _shortcutInternals } from '../src/renderer/excel-shortcuts'

function shape(
  startRow: number,
  startColumn: number,
  endRow: number,
  endColumn: number,
  rangeType?: RANGE_TYPE,
): SelectionShape {
  return {
    range: {
      startRow,
      startColumn,
      endRow,
      endColumn,
      ...(rangeType === undefined ? {} : { rangeType }),
    },
    rowCount: 100,
    columnCount: 26,
  }
}

describe('decideCellsAction', () => {
  it('acts on whole rows without asking', () => {
    expect(decideCellsAction(shape(3, 0, 4, 25, RANGE_TYPE.ROW))).toEqual({
      kind: 'direct',
      choice: 'entire-row',
    })
    expect(decideCellsAction(shape(3, 0, 3, 25))).toEqual({ kind: 'direct', choice: 'entire-row' })
  })

  it('acts on whole columns without asking', () => {
    expect(decideCellsAction(shape(0, 2, 99, 2, RANGE_TYPE.COLUMN))).toEqual({
      kind: 'direct',
      choice: 'entire-column',
    })
    expect(decideCellsAction(shape(0, 2, 99, 3))).toEqual({
      kind: 'direct',
      choice: 'entire-column',
    })
  })

  it('treats select-all as rows', () => {
    expect(decideCellsAction(shape(0, 0, 99, 25, RANGE_TYPE.ALL))).toEqual({
      kind: 'direct',
      choice: 'entire-row',
    })
  })

  it('opens the dialog for any other range', () => {
    expect(decideCellsAction(shape(1, 1, 1, 1))).toEqual({ kind: 'dialog' })
    expect(decideCellsAction(shape(2, 0, 5, 24))).toEqual({ kind: 'dialog' })
    expect(decideCellsAction(shape(1, 3, 99, 3))).toEqual({ kind: 'dialog' })
  })
})

describe('cellsCommandId', () => {
  it('maps insert choices to the Univer commands', () => {
    expect(cellsCommandId('insert', 'shift-horizontal')).toBe(INSERT_RANGE_RIGHT_COMMAND)
    expect(cellsCommandId('insert', 'shift-vertical')).toBe(INSERT_RANGE_DOWN_COMMAND)
    expect(cellsCommandId('insert', 'entire-row')).toBe(INSERT_ROWS_COMMAND)
    expect(cellsCommandId('insert', 'entire-column')).toBe(INSERT_COLS_COMMAND)
  })

  it('maps delete choices to the Univer commands', () => {
    expect(cellsCommandId('delete', 'shift-horizontal')).toBe(DELETE_RANGE_LEFT_COMMAND)
    expect(cellsCommandId('delete', 'shift-vertical')).toBe(DELETE_RANGE_UP_COMMAND)
    expect(cellsCommandId('delete', 'entire-row')).toBe(REMOVE_ROWS_COMMAND)
    expect(cellsCommandId('delete', 'entire-column')).toBe(REMOVE_COLS_COMMAND)
  })

  it('defaults to the vertical shift like Excel', () => {
    expect(DEFAULT_CELLS_CHOICE).toBe('shift-vertical')
  })
})

describe('cellsCommandParams', () => {
  const range = { startRow: 2, startColumn: 1, endRow: 4, endColumn: 2 }

  it('passes the selected line count to the insert-before commands', () => {
    expect(cellsCommandParams('insert', 'entire-row', range)).toEqual({ value: 3 })
    expect(cellsCommandParams('insert', 'entire-column', range)).toEqual({ value: 2 })
  })

  it('lets the other commands read the selection themselves', () => {
    expect(cellsCommandParams('insert', 'shift-vertical', range)).toBeUndefined()
    expect(cellsCommandParams('delete', 'entire-row', range)).toBeUndefined()
  })
})

describe('numpad keycodes', () => {
  it('match the browser values the shortcut dispatcher sees', () => {
    expect([_shortcutInternals.KEY_NUMPAD_ADD, _shortcutInternals.KEY_NUMPAD_SUBTRACT]).toEqual([
      107, 109,
    ])
  })
})
