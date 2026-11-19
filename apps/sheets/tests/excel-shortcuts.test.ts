/**
 * Excel-standard shortcuts Univer doesn't ship: worksheet-tab switching,
 * Ctrl+Home/End, Home, whole row/column selection.
 */
import { describe, expect, it } from 'vitest'
import { Styles, Worksheet } from '@univerjs/core'
import type { IWorksheetData } from '@univerjs/core'
import {
  _shortcutInternals,
  currentRegion,
  extendedRange,
  farCorner,
} from '../src/renderer/excel-shortcuts'

const { unfrozenOrigin, KEY_PAGE_UP, KEY_PAGE_DOWN, KEY_END, KEY_HOME } = _shortcutInternals

function makeSheet(freeze?: IWorksheetData['freeze']): Worksheet {
  const snapshot: Partial<IWorksheetData> = {
    id: 's1',
    name: 'Test',
    rowCount: 100,
    columnCount: 26,
    cellData: {},
    mergeData: [],
    ...(freeze ? { freeze } : {}),
  }
  return new Worksheet('unit', snapshot as IWorksheetData, new Styles())
}

describe('browser keycodes missing from Univer KeyCode enum', () => {
  it('uses the standard values the shortcut dispatcher matches on', () => {
    expect([KEY_PAGE_UP, KEY_PAGE_DOWN, KEY_END, KEY_HOME]).toEqual([33, 34, 35, 36])
  })
})

describe('unfrozenOrigin (Ctrl+Home target)', () => {
  it('is A1 without frozen panes', () => {
    expect(unfrozenOrigin(makeSheet())).toEqual({ row: 0, column: 0 })
  })

  it('is the first cell below/right of the freeze', () => {
    expect(
      unfrozenOrigin(makeSheet({ xSplit: 1, ySplit: 2, startRow: 2, startColumn: 1 })),
    ).toEqual({ row: 2, column: 1 })
  })

  it('ignores a freeze on one axis only for the other axis', () => {
    expect(
      unfrozenOrigin(makeSheet({ xSplit: 0, ySplit: 3, startRow: 3, startColumn: -1 })),
    ).toEqual({ row: 3, column: 0 })
  })
})

describe('unfrozenOrigin skips hidden lines', () => {
  it('lands on the first visible column when column A is hidden', () => {
    const snapshot: Partial<IWorksheetData> = {
      id: 's1',
      name: 'Test',
      rowCount: 100,
      columnCount: 26,
      cellData: {},
      mergeData: [],
      columnData: { 0: { hd: 1 } },
    }
    const sheet = new Worksheet('unit', snapshot as IWorksheetData, new Styles())
    expect(unfrozenOrigin(sheet)).toEqual({ row: 0, column: 1 })
  })

  it('lands below hidden top rows', () => {
    const snapshot: Partial<IWorksheetData> = {
      id: 's1',
      name: 'Test',
      rowCount: 100,
      columnCount: 26,
      cellData: {},
      mergeData: [],
      rowData: { 0: { hd: 1 }, 1: { hd: 1 } },
    }
    const sheet = new Worksheet('unit', snapshot as IWorksheetData, new Styles())
    expect(unfrozenOrigin(sheet)).toEqual({ row: 2, column: 0 })
  })
})

describe('extendedRange (Shift+Home / Ctrl+Shift+Home / Ctrl+Shift+End)', () => {
  const range = { startRow: 2, endRow: 4, startColumn: 3, endColumn: 5 }

  it('moves the corner opposite the anchor and keeps the other axis', () => {
    expect(extendedRange(range, { row: 2, column: 3 }, { row: null, column: 0 })).toEqual({
      startRow: 2,
      endRow: 4,
      startColumn: 0,
      endColumn: 3,
    })
  })

  it('keeps the whole span of an untouched axis when the anchor sits inside it', () => {
    // Ctrl+Shift+* leaves the active cell mid-block; Shift+Home must not drop rows
    expect(extendedRange(range, { row: 3, column: 4 }, { row: null, column: 0 })).toEqual({
      startRow: 2,
      endRow: 4,
      startColumn: 0,
      endColumn: 4,
    })
  })

  it('collapses past the anchor when the target lies on its far side', () => {
    expect(extendedRange(range, { row: 4, column: 5 }, { row: 0, column: 0 })).toEqual({
      startRow: 0,
      endRow: 4,
      startColumn: 0,
      endColumn: 5,
    })
    expect(extendedRange(range, { row: 2, column: 3 }, { row: 9, column: 9 })).toEqual({
      startRow: 2,
      endRow: 9,
      startColumn: 3,
      endColumn: 9,
    })
  })
})

describe('farCorner (cell revealed after extending)', () => {
  const range = { startRow: 0, endRow: 9, startColumn: 0, endColumn: 4 }
  const both = { row: 0, column: 0 }
  it('is the corner diagonally away from the anchor', () => {
    expect(farCorner(range, { row: 0, column: 0 }, both)).toEqual({ row: 9, column: 4 })
    expect(farCorner(range, { row: 9, column: 4 }, both)).toEqual({ row: 0, column: 0 })
    expect(farCorner(range, { row: 9, column: 0 }, both)).toEqual({ row: 0, column: 4 })
  })
  it('stays on the anchor row when only the column moved (Shift+Home)', () => {
    expect(farCorner(range, { row: 5, column: 4 }, { row: null, column: 0 })).toEqual({
      row: 5,
      column: 0,
    })
  })
})

describe('currentRegion (Ctrl+Shift+*)', () => {
  const filled = new Set(['1,1', '1,2', '2,1', '2,2', '3,2', '6,6'])
  const hasValue = (row: number, column: number): boolean => filled.has(`${row},${column}`)

  it('grows to the block bordered by empty rows and columns', () => {
    expect(
      currentRegion({ startRow: 2, endRow: 2, startColumn: 2, endColumn: 2 }, hasValue, 100, 26),
    ).toEqual({ startRow: 1, endRow: 3, startColumn: 1, endColumn: 2 })
  })

  it('stays on an isolated empty cell', () => {
    expect(
      currentRegion(
        { startRow: 10, endRow: 10, startColumn: 10, endColumn: 10 },
        hasValue,
        100,
        26,
      ),
    ).toEqual({ startRow: 10, endRow: 10, startColumn: 10, endColumn: 10 })
  })

  it('stops at the sheet edges', () => {
    const corner = (row: number, column: number): boolean => row <= 1 && column <= 1
    expect(
      currentRegion({ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, corner, 2, 2),
    ).toEqual({ startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 })
  })
})
