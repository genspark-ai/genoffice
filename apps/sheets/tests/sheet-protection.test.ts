import { describe, expect, it } from 'vitest'

import {
  CELL_LOCKED_KEY,
  cellIsLocked,
  commandGuardKind,
  fileSheetProtection,
  mergeSheetProtectionAllow,
  rangesInParams,
  rangesTouchLockedCell,
  setRangeValuesChangesContent,
} from '../src/renderer/sheet-protection'

describe('sheet protection command mapping', () => {
  it('maps Univer commands onto the Excel permission flags', () => {
    expect(commandGuardKind('sheet.operation.set-cell-edit-visible')).toBe('edit')
    expect(commandGuardKind('sheet.command.paste')).toBe('edit')
    expect(commandGuardKind('sheet.command.set-range-bold')).toBe('formatCells')
    expect(commandGuardKind('sheet.command.set-worksheet-col-width')).toBe('formatColumns')
    expect(commandGuardKind('sheet.command.set-row-height')).toBe('formatRows')
    expect(commandGuardKind('sheet.command.insert-col-before')).toBe('insertColumns')
    expect(commandGuardKind('sheet.command.insert-row-after')).toBe('insertRows')
    expect(commandGuardKind('sheet.command.remove-col-confirm')).toBe('deleteColumns')
    expect(commandGuardKind('sheet.command.remove-row-confirm')).toBe('deleteRows')
    expect(commandGuardKind('sheet.command.sort-range-asc')).toBe('sort')
    expect(commandGuardKind('sheet.command.smart-toggle-filter')).toBe('autoFilter')
    expect(commandGuardKind('sheet.command.set-worksheet-activate')).toBeUndefined()
    expect(commandGuardKind('sheet.command.scroll-view')).toBeUndefined()
  })

  it('tells content writes from style-only set-range-values payloads', () => {
    expect(setRangeValuesChangesContent({ v: 1 })).toBe(true)
    expect(setRangeValuesChangesContent({ s: { bl: 1 } })).toBe(false)
    expect(setRangeValuesChangesContent({ custom: { [CELL_LOCKED_KEY]: false } })).toBe(false)
    expect(setRangeValuesChangesContent({ 0: { 0: { s: { bl: 1 } }, 1: { f: '=A1' } } })).toBe(true)
    expect(setRangeValuesChangesContent({ 0: { 0: null } })).toBe(true)
    expect(setRangeValuesChangesContent([[{ s: null }]])).toBe(false)
  })

  it('collects range-shaped params', () => {
    const range = { startRow: 1, endRow: 2, startColumn: 0, endColumn: 0 }
    expect(rangesInParams({ unitId: 'u', targetRange: range, ranges: [range] })).toEqual([
      range,
      range,
    ])
    expect(rangesInParams({ visible: true })).toEqual([])
  })
})

describe('locked cell detection', () => {
  const unlocked = new Set(['0:1', '2:2'])
  const lookup = (defaultLocked = true, allowEditRefs: string[] = []) => ({
    getCell: (row: number, column: number) =>
      unlocked.has(`${row}:${column}`) ? { custom: { [CELL_LOCKED_KEY]: false } } : null,
    defaultLocked,
    allowEditRefs,
    maxRow: 99,
    maxColumn: 25,
    residentRows: null,
    frozenRows: 0,
  })
  const cell = (row: number, column: number) => ({
    startRow: row,
    endRow: row,
    startColumn: column,
    endColumn: column,
  })

  it('treats cells without a flag as locked unless the Normal xf unlocks', () => {
    expect(cellIsLocked(null, true)).toBe(true)
    expect(cellIsLocked({ custom: { [CELL_LOCKED_KEY]: false } }, true)).toBe(false)
    expect(cellIsLocked({ custom: {} }, false)).toBe(false)
    expect(rangesTouchLockedCell([cell(0, 1)], lookup())).toBe(false)
    expect(rangesTouchLockedCell([cell(0, 0)], lookup())).toBe(true)
    expect(rangesTouchLockedCell([cell(0, 0)], lookup(false))).toBe(false)
  })

  it('fails closed for rows outside the resident window', () => {
    const partial = { ...lookup(false), residentRows: { startRow: 20, endRow: 29 }, frozenRows: 2 }
    expect(rangesTouchLockedCell([cell(25, 0)], partial)).toBe(false)
    expect(rangesTouchLockedCell([cell(50, 0)], partial)).toBe(true)
    // Frozen rows stay installed while the viewport streams elsewhere.
    expect(rangesTouchLockedCell([cell(1, 0)], partial)).toBe(false)
    expect(rangesTouchLockedCell([cell(2, 0)], partial)).toBe(true)
  })

  it('blocks a range as soon as one cell in it is locked', () => {
    expect(
      rangesTouchLockedCell([{ startRow: 0, endRow: 0, startColumn: 1, endColumn: 2 }], lookup()),
    ).toBe(true)
  })

  it('keeps allow-edit ranges editable and clips whole-column selections', () => {
    expect(rangesTouchLockedCell([cell(5, 3)], lookup(true, ['D6:D10 F1']))).toBe(false)
    expect(rangesTouchLockedCell([cell(5, 3)], lookup(true, ['$D$6:$D$10']))).toBe(false)
    expect(rangesTouchLockedCell([cell(5, 3)], lookup(true, ['D7:D10']))).toBe(true)
    const column = { startRow: 0, endRow: 1_048_575, startColumn: 3, endColumn: 3 }
    expect(rangesTouchLockedCell([column], lookup(true, ['D1:D100']))).toBe(false)
  })
})

describe('protection records', () => {
  it('normalizes sidecar records and dialog partials', () => {
    expect(fileSheetProtection(null).protected).toBe(false)
    const file = fileSheetProtection({ protected: true, hasPassword: true })
    expect(file.allow.selectLockedCells).toBe(true)
    expect(file.password).toBeNull()
    const allow = mergeSheetProtectionAllow({ sort: true, formatCells: undefined })
    expect(allow.sort).toBe(true)
    expect(allow.formatCells).toBe(false)
    expect(allow.selectUnlockedCells).toBe(true)
  })
})
