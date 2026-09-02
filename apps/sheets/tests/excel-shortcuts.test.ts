/**
 * Excel-standard shortcuts Univer doesn't ship: worksheet-tab switching,
 * Ctrl+Home/End, Home, whole row/column selection (#genoffice-6 task #1).
 */
import { describe, expect, it } from 'vitest'
import { Styles, Worksheet } from '@univerjs/core'
import type { IWorksheetData } from '@univerjs/core'
import { _shortcutInternals } from '../src/renderer/excel-shortcuts'

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
