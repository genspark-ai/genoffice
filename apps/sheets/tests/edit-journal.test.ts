import { describe, expect, it } from 'vitest'

import {
  createEditJournal,
  isSheetRemoved,
  recordFilterChange,
  fromNeutralStyle,
  journalEntriesInRange,
  journalSize,
  recordPageSetup,
  recordProtectedRangesChange,
  recordSetNumfmt,
  recordSetRangeValues,
  recordThemeColors,
  recordThemeFonts,
  recordWorkbookProtection,
  recordSheetDuplicate,
  recordSheetHidden,
  recordSheetInsert,
  recordSheetOrderChange,
  recordSheetRemove,
  recordSheetRename,
  recordChartEdit,
  recordPivotAdd,
  recordStructuralOp,
  recordTableAdd,
  recordVisualAdd,
  toNeutralStyle,
  toRecalcUserInput,
  toSaveChartEdits,
  toSaveEdits,
  toSavePivotAdds,
  toSaveTableAdds,
  toSaveSheetOps,
  toSaveStructuralOps,
} from '../src/renderer/edit-journal'

describe('recordTableAdd', () => {
  const table = {
    sheetId: 'sheet-1',
    area: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 1 },
    name: 'Users',
    columnNames: ['Country', 'Users'],
    bandedRows: true,
  }

  it('counts toward the pending-edit badge and reaches the save payload', () => {
    const journal = createEditJournal()
    recordTableAdd(journal, table)
    expect(journalSize(journal)).toBe(1)
    expect(toSaveTableAdds(journal)).toEqual([table])
  })

  it('drops tables on removed sheets', () => {
    const journal = createEditJournal()
    recordTableAdd(journal, table)
    recordSheetRemove(journal, 'sheet-1')
    expect(journalSize(journal)).toBe(1) // the sheet removal itself
    expect(toSaveTableAdds(journal)).toEqual([])
  })
})

describe('recordPivotAdd', () => {
  const pivot = {
    sheetId: 'sheet-2',
    sourceSheetId: 'sheet-1',
    sourceArea: { startRow: 0, startColumn: 0, endRow: 99, endColumn: 3 },
    location: { startRow: 0, startColumn: 0, endRow: 5, endColumn: 1 },
    name: 'Pivot1',
    fieldNames: ['Region', 'Product', 'Quantity', 'Amount'],
    rowFieldIndices: [0],
    rowItems: ['East', 'South'],
    values: [{ fieldIndex: 3, agg: 'sum' as const }],
  }

  it('counts toward the badge and reaches the save payload', () => {
    const journal = createEditJournal()
    recordPivotAdd(journal, pivot)
    expect(journalSize(journal)).toBe(1)
    expect(toSavePivotAdds(journal)).toEqual([pivot])
  })

  it('drops pivots when the source or output sheet was removed', () => {
    const journal = createEditJournal()
    recordPivotAdd(journal, pivot)
    recordSheetRemove(journal, 'sheet-1')
    expect(toSavePivotAdds(journal)).toEqual([])
  })
})

describe('recordSetRangeValues', () => {
  it('records scalar values from a mutation matrix', () => {
    const journal = createEditJournal()
    const recorded = recordSetRangeValues(journal, 'sheet-1', {
      2: { 3: { v: 42 }, 4: { v: 'text' } },
      5: { 0: { v: true } },
    })
    expect(recorded).toHaveLength(3)
    expect(journalSize(journal)).toBe(3)
    expect(journal.cells.get('sheet-1')?.get('2:3')).toEqual({
      row: 2,
      column: 3,
      hasValue: true,
      value: 42,
    })
  })

  it('normalizes formulas to a leading equals sign', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { f: 'SUM(A1:A2)' } } })
    recordSetRangeValues(journal, 'sheet-1', { 0: { 1: { f: '=MAX(B1:B2)' } } })
    expect(journal.cells.get('sheet-1')?.get('0:0')?.formula).toBe('=SUM(A1:A2)')
    expect(journal.cells.get('sheet-1')?.get('0:1')?.formula).toBe('=MAX(B1:B2)')
  })

  it('treats null and empty-value cells as content clears', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 1: { 1: null, 3: { v: null } } })
    // A whole-cell null (Clear All) also resets formatting; a bare value
    // clear keeps the cell's style.
    expect(journal.cells.get('sheet-1')?.get('1:1')).toEqual({
      row: 1,
      column: 1,
      hasValue: true,
      value: null,
      styleReset: true,
    })
    expect(journal.cells.get('sheet-1')?.get('1:3')).toEqual({
      row: 1,
      column: 3,
      hasValue: true,
      value: null,
    })
  })

  it('records Clear Formats (s: null) as a style reset that drops earlier deltas', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { s: { bl: 1 } } } })
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { s: null } } })
    expect(journal.cells.get('sheet-1')?.get('0:0')).toEqual({
      row: 0,
      column: 0,
      hasValue: false,
      value: null,
      styleReset: true,
    })
  })

  it('stacks a delta arriving after a reset on top of the default style', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { s: null } } })
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { s: { it: 1 } } } })
    const entry = journal.cells.get('sheet-1')?.get('0:0')
    expect(entry?.styleReset).toBe(true)
    expect(entry?.style).toEqual({ italic: true })
  })

  it('flattens rich-text edits to plain text', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', {
      0: { 0: { p: { body: { dataStream: 'styled text\r\n' } } } },
    })
    expect(journal.cells.get('sheet-1')?.get('0:0')?.value).toBe('styled text')
  })

  it('captures a style delta without marking the value as edited', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { s: { bl: 1 } } } })
    expect(journal.cells.get('sheet-1')?.get('0:0')).toEqual({
      row: 0,
      column: 0,
      hasValue: false,
      value: null,
      style: { bold: true },
    })
  })

  it('merges successive style deltas and value edits on one cell', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { s: { bl: 1 } } } })
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { s: { cl: { rgb: '#c00000' } } } } })
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { v: 7 } } })
    expect(journal.cells.get('sheet-1')?.get('0:0')).toEqual({
      row: 0,
      column: 0,
      hasValue: true,
      value: 7,
      style: { bold: true, fontColor: '#C00000' },
    })
    expect(journalSize(journal)).toBe(1)
  })

  it('keeps a journaled formula when the engine writes its cached result', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { f: '=SUM(1,2,3)' } } })
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { v: 6 } } })
    expect(journal.cells.get('sheet-1')?.get('0:0')?.formula).toBe('=SUM(1,2,3)')
  })

  it('clears a journaled formula on an explicit editor overwrite (f: null)', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { f: '=SUM(1,2,3)' } } })
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { v: 7, f: null, si: null } } })
    expect(journal.cells.get('sheet-1')?.get('0:0')).toEqual({
      row: 0,
      column: 0,
      hasValue: true,
      value: 7,
    })
  })

  it('ignores malformed keys and non-finite numbers', () => {
    const journal = createEditJournal()
    const recorded = recordSetRangeValues(journal, 'sheet-1', {
      'not-a-row': { 0: { v: 1 } },
      0: { 'not-a-column': { v: 1 }, 1: { v: Number.NaN } },
    })
    expect(recorded).toHaveLength(0)
    expect(journalSize(journal)).toBe(0)
  })
})

describe('recordSetNumfmt', () => {
  it('expands pattern ranges into per-cell style entries', () => {
    const journal = createEditJournal()
    const recorded = recordSetNumfmt(journal, 'sheet-1', {
      refMap: { 1: { pattern: '0.00%' } },
      values: { 1: { ranges: [{ startRow: 2, endRow: 3, startColumn: 1, endColumn: 1 }] } },
    })
    expect(recorded).toHaveLength(2)
    expect(journal.cells.get('sheet-1')?.get('2:1')?.style).toEqual({ numberFormat: '0.00%' })
    expect(journal.cells.get('sheet-1')?.get('3:1')?.hasValue).toBe(false)
  })

  it('rejects unbounded ranges', () => {
    const journal = createEditJournal()
    const recorded = recordSetNumfmt(journal, 'sheet-1', {
      refMap: { 1: { pattern: '0.00' } },
      values: {
        1: {
          ranges: [{ startRow: 0, endRow: Number.MAX_SAFE_INTEGER, startColumn: 0, endColumn: 0 }],
        },
      },
    })
    expect(recorded).toHaveLength(0)
  })
})

describe('style conversion', () => {
  it('maps Univer deltas to the neutral wire format', () => {
    expect(
      toNeutralStyle({
        bl: 1,
        it: null,
        ul: { s: 1 },
        st: null,
        ff: 'Arial',
        fs: 14,
        cl: { rgb: '#c00000' },
        bg: { rgb: 'FFF2CC' },
        ht: 2,
        vt: 3,
        tb: 3,
        n: { pattern: '#,##0' },
      }),
    ).toEqual({
      bold: true,
      italic: false,
      underline: true,
      underlineStyle: 'single',
      strikethrough: false,
      fontFamily: 'Arial',
      fontSize: 14,
      fontColor: '#C00000',
      fillColor: '#FFF2CC',
      horizontalAlignment: 'center',
      verticalAlignment: 'bottom',
      wrapText: true,
      numberFormat: '#,##0',
    })
  })

  it('maps text rotation and double underline both ways', () => {
    expect(toNeutralStyle({ tr: { a: 45 } })).toEqual({ textRotation: 45 })
    expect(toNeutralStyle({ tr: { a: -45 } })).toEqual({ textRotation: 135 })
    expect(toNeutralStyle({ tr: { a: 0, v: 1 } })).toEqual({ textRotation: 255 })
    expect(toNeutralStyle({ tr: null })).toEqual({ textRotation: 0 })
    expect(toNeutralStyle({ ul: { s: 1, t: 10 } })).toEqual({
      underline: true,
      underlineStyle: 'double',
    })
    expect(fromNeutralStyle({ textRotation: 135 })).toEqual({ tr: { a: -45 } })
    expect(fromNeutralStyle({ textRotation: 255 })).toEqual({ tr: { a: 0, v: 1 } })
    expect(fromNeutralStyle({ textRotation: 0 })).toEqual({ tr: null })
    expect(fromNeutralStyle({ underline: true, underlineStyle: 'double' })).toEqual({
      ul: { s: 1, t: 10 },
    })
  })

  it('maps indent to left padding both ways', () => {
    expect(toNeutralStyle({ pd: { l: 16 } })).toEqual({ indent: 2 })
    expect(toNeutralStyle({ pd: null })).toEqual({ indent: 0 })
    // Foreign padding rounds to the nearest step and clamps to OOXML's 250.
    expect(toNeutralStyle({ pd: { l: 13 } })).toEqual({ indent: 2 })
    expect(toNeutralStyle({ pd: { l: 99_999 } })).toEqual({ indent: 250 })
    expect(fromNeutralStyle({ indent: 2 })).toEqual({ pd: { l: 16 } })
    expect(fromNeutralStyle({ indent: 0 })).toEqual({ pd: null })
  })

  it('round-trips through fromNeutralStyle', () => {
    const neutral = toNeutralStyle({ bl: 1, cl: { rgb: '#C00000' }, ht: 3, tb: 3 })
    expect(neutral).toBeDefined()
    expect(fromNeutralStyle(neutral ?? {})).toEqual({
      bl: 1,
      cl: { rgb: '#C00000' },
      ht: 3,
      tb: 3,
    })
  })

  it('returns undefined for an empty or unknown delta', () => {
    expect(toNeutralStyle({})).toBeUndefined()
    expect(toNeutralStyle({ unknown: true })).toBeUndefined()
  })
})

describe('journalEntriesInRange', () => {
  it('returns only entries inside the range for the right sheet', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 10: { 2: { v: 'in' } }, 99: { 2: { v: 'out' } } })
    recordSetRangeValues(journal, 'sheet-2', { 10: { 2: { v: 'other sheet' } } })
    const entries = journalEntriesInRange(journal, 'sheet-1', {
      startRow: 0,
      endRow: 50,
      startColumn: 0,
      endColumn: 5,
    })
    expect(entries).toEqual([{ row: 10, column: 2, hasValue: true, value: 'in' }])
  })
})

describe('recordStructuralOp', () => {
  it('shifts existing cell entries below an inserted row', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 2: { 0: { v: 'above' } }, 8: { 0: { v: 'below' } } })
    recordStructuralOp(journal, 'sheet-1', { kind: 'insert-rows', index: 5, count: 2 })
    expect(journal.cells.get('sheet-1')?.get('2:0')?.value).toBe('above')
    expect(journal.cells.get('sheet-1')?.get('10:0')?.value).toBe('below')
    expect(journal.cells.get('sheet-1')?.has('8:0')).toBe(false)
  })

  it('drops entries inside a removed range and shifts the rest', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', {
      0: { 1: { v: 'left' }, 2: { v: 'gone' }, 4: { v: 'shifts' } },
    })
    recordStructuralOp(journal, 'sheet-1', { kind: 'remove-cols', index: 2, count: 2 })
    expect(journal.cells.get('sheet-1')?.get('0:1')?.value).toBe('left')
    expect(journal.cells.get('sheet-1')?.get('0:2')?.value).toBe('shifts')
    expect(journalSize(journal)).toBe(3)
  })

  it('an exact inverse remove (undo of a just-recorded insert) cancels the pair', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 8: { 0: { v: 'below' } } })
    recordStructuralOp(journal, 'sheet-1', { kind: 'insert-rows', index: 5, count: 2 })
    expect(journal.cells.get('sheet-1')?.has('10:0')).toBe(true)
    recordStructuralOp(journal, 'sheet-1', { kind: 'remove-rows', index: 5, count: 2 })
    // net no-op: no structural ops left, the cell entry back at its original spot
    expect(journal.structuralOps.has('sheet-1')).toBe(false)
    expect(journal.cells.get('sheet-1')?.get('8:0')?.value).toBe('below')

    // a remove that is NOT the inverse of the last op still stacks
    recordStructuralOp(journal, 'sheet-1', { kind: 'insert-rows', index: 5, count: 2 })
    recordStructuralOp(journal, 'sheet-1', { kind: 'remove-rows', index: 3, count: 2 })
    expect(journal.structuralOps.get('sheet-1')).toHaveLength(2)
  })

  it('records sizing and visibility ops without shifting cell entries', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 5: { 2: { v: 'keep' } } })
    recordStructuralOp(journal, 'sheet-1', {
      kind: 'set-row-size',
      start: 0,
      end: 9,
      size: 24,
    })
    recordStructuralOp(journal, 'sheet-1', {
      kind: 'set-cols-hidden',
      start: 1,
      end: 1,
      hidden: true,
    })
    expect(journal.cells.get('sheet-1')?.get('5:2')).toEqual({
      row: 5,
      column: 2,
      hasValue: true,
      value: 'keep',
    })
    expect(toSaveStructuralOps(journal)).toEqual([
      { sheetId: 'sheet-1', kind: 'set-row-size', start: 0, end: 9, size: 24 },
      { sheetId: 'sheet-1', kind: 'set-cols-hidden', start: 1, end: 1, hidden: true },
    ])
  })

  it('records outline ops without shifting cell entries', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 5: { 2: { v: 'keep' } } })
    recordStructuralOp(journal, 'sheet-1', {
      kind: 'set-rows-outline',
      start: 1,
      end: 4,
      level: 1,
    })
    recordStructuralOp(journal, 'sheet-1', {
      kind: 'set-cols-outline',
      start: 5,
      end: 5,
      level: 0,
      collapsed: true,
    })
    expect(journal.cells.get('sheet-1')?.get('5:2')).toEqual({
      row: 5,
      column: 2,
      hasValue: true,
      value: 'keep',
    })
    expect(toSaveStructuralOps(journal)).toEqual([
      { sheetId: 'sheet-1', kind: 'set-rows-outline', start: 1, end: 4, level: 1 },
      { sheetId: 'sheet-1', kind: 'set-cols-outline', start: 5, end: 5, level: 0, collapsed: true },
    ])
  })

  it('records merges without shifting cell entries', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 5: { 0: { v: 'stays' } } })
    recordStructuralOp(journal, 'sheet-1', {
      kind: 'merge-cells',
      range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 2 },
    })
    expect(journal.cells.get('sheet-1')?.get('5:0')?.value).toBe('stays')
    expect(toSaveStructuralOps(journal)).toEqual([
      {
        sheetId: 'sheet-1',
        kind: 'merge-cells',
        range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 2 },
      },
    ])
  })

  it('counts ops in journalSize and flattens them per sheet in order', () => {
    const journal = createEditJournal()
    recordStructuralOp(journal, 'sheet-1', { kind: 'insert-rows', index: 1, count: 1 })
    recordStructuralOp(journal, 'sheet-1', { kind: 'remove-cols', index: 0, count: 2 })
    expect(journalSize(journal)).toBe(2)
    expect(toSaveStructuralOps(journal)).toEqual([
      { sheetId: 'sheet-1', kind: 'insert-rows', index: 1, count: 1 },
      { sheetId: 'sheet-1', kind: 'remove-cols', index: 0, count: 2 },
    ])
  })
})

describe('toSaveEdits', () => {
  it('flattens the journal into IPC-ready edits', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { v: 'a' }, 1: { f: '=SUM(A1)' } } })
    recordSetRangeValues(journal, 'sheet-2', { 3: { 4: { s: { bl: 1 } } } })
    const edits = toSaveEdits(journal)
    expect(edits).toHaveLength(3)
    expect(edits).toContainEqual({
      sheetId: 'sheet-1',
      row: 0,
      column: 0,
      writeValue: true,
      value: 'a',
    })
    expect(edits).toContainEqual({
      sheetId: 'sheet-1',
      row: 0,
      column: 1,
      writeValue: true,
      value: null,
      formula: '=SUM(A1)',
    })
    expect(edits).toContainEqual({
      sheetId: 'sheet-2',
      row: 3,
      column: 4,
      writeValue: false,
      value: null,
      style: { bold: true },
    })
  })
})

describe('sheet lifecycle journal', () => {
  it('an added sheet renamed later saves as a single add with the final name', () => {
    const journal = createEditJournal()
    recordSheetInsert(journal, 'u-new', 'Sheet4')
    recordSheetRename(journal, 'u-new', 'Budget', undefined)
    expect(toSaveSheetOps(journal)).toEqual([
      { kind: 'add-sheet', sheetId: 'u-new', name: 'Budget' },
    ])
    expect(journalSize(journal)).toBe(1)
  })

  it('adding then removing a sheet cancels out', () => {
    const journal = createEditJournal()
    recordSheetInsert(journal, 'u-new', 'Sheet4')
    recordSetRangeValues(journal, 'u-new', { 0: { 0: { v: 1 } } })
    recordSheetRemove(journal, 'u-new')
    expect(toSaveSheetOps(journal)).toEqual([])
    expect(toSaveEdits(journal)).toEqual([])
    expect(journalSize(journal)).toBe(0)
  })

  it('renaming back to the original name drops the op', () => {
    const journal = createEditJournal()
    recordSheetRename(journal, 'sheet-1', 'Renamed', 'Data')
    expect(toSaveSheetOps(journal)).toEqual([
      { kind: 'rename-sheet', sheetId: 'sheet-1', newName: 'Renamed' },
    ])
    recordSheetRename(journal, 'sheet-1', 'Data', 'Data')
    expect(toSaveSheetOps(journal)).toEqual([])
  })

  it('removing an original sheet keeps its entries but excludes them from the payload', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-2', { 0: { 0: { v: 9 } } })
    recordStructuralOp(journal, 'sheet-2', { kind: 'insert-rows', index: 0, count: 1 })
    recordSheetRename(journal, 'sheet-2', 'Gone', 'Other')
    recordSheetRemove(journal, 'sheet-2')
    expect(toSaveSheetOps(journal)).toEqual([{ kind: 'remove-sheet', sheetId: 'sheet-2' }])
    expect(toSaveEdits(journal)).toEqual([])
    expect(toSaveStructuralOps(journal)).toEqual([])
    // Undo re-inserts the sheet: the mark clears and the entries are live again.
    recordSheetInsert(journal, 'sheet-2', 'Other')
    expect(toSaveSheetOps(journal)).toEqual([
      { kind: 'rename-sheet', sheetId: 'sheet-2', newName: 'Gone' },
    ])
    expect(toSaveEdits(journal)).toHaveLength(1)
    expect(toSaveStructuralOps(journal)).toHaveLength(1)
  })
})

describe('sheet duplicate journal', () => {
  it('copies pending edits and structural ops into the duplicate at copy time', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { v: 1 } } })
    recordStructuralOp(journal, 'sheet-1', { kind: 'insert-rows', index: 0, count: 1 })
    recordSheetDuplicate(journal, 'u-copy', 'Data Copy', 'sheet-1')
    // Edits to the source after the copy must not leak into it.
    recordSetRangeValues(journal, 'sheet-1', { 5: { 5: { v: 2 } } })
    expect(toSaveSheetOps(journal)).toEqual([
      { kind: 'duplicate-sheet', sheetId: 'u-copy', name: 'Data Copy', sourceSheetId: 'sheet-1' },
    ])
    const copyEdits = toSaveEdits(journal).filter((edit) => edit.sheetId === 'u-copy')
    expect(copyEdits).toHaveLength(1)
    expect(copyEdits[0]).toMatchObject({ row: 1, column: 0, value: 1 })
    expect(toSaveStructuralOps(journal).filter((op) => op.sheetId === 'u-copy')).toEqual([
      { sheetId: 'u-copy', kind: 'insert-rows', index: 0, count: 1 },
    ])
    expect(toSaveEdits(journal).filter((edit) => edit.sheetId === 'sheet-1')).toHaveLength(2)
  })

  it('a duplicate of a session-added sheet saves as a plain add', () => {
    const journal = createEditJournal()
    recordSheetInsert(journal, 'u-new', 'Fresh')
    recordSetRangeValues(journal, 'u-new', { 0: { 0: { v: 'seed' } } })
    recordSheetDuplicate(journal, 'u-copy', 'Fresh Copy', 'u-new')
    expect(toSaveSheetOps(journal)).toEqual([
      { kind: 'add-sheet', sheetId: 'u-new', name: 'Fresh' },
      { kind: 'add-sheet', sheetId: 'u-copy', name: 'Fresh Copy' },
    ])
    expect(toSaveEdits(journal).filter((edit) => edit.sheetId === 'u-copy')).toHaveLength(1)
  })

  it('a duplicate of a duplicate resolves to the original file sheet', () => {
    const journal = createEditJournal()
    recordSheetDuplicate(journal, 'u-copy', 'Data Copy', 'sheet-1')
    recordSheetDuplicate(journal, 'u-copy2', 'Data Copy 2', 'u-copy')
    expect(toSaveSheetOps(journal)).toEqual([
      { kind: 'duplicate-sheet', sheetId: 'u-copy', name: 'Data Copy', sourceSheetId: 'sheet-1' },
      {
        kind: 'duplicate-sheet',
        sheetId: 'u-copy2',
        name: 'Data Copy 2',
        sourceSheetId: 'sheet-1',
      },
    ])
  })

  it('removing a duplicate cancels it; undo restores it with its source intact', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { v: 1 } } })
    recordSheetDuplicate(journal, 'u-copy', 'Data Copy', 'sheet-1')
    recordSheetRemove(journal, 'u-copy')
    expect(toSaveSheetOps(journal)).toEqual([])
    expect(toSaveEdits(journal).filter((edit) => edit.sheetId === 'u-copy')).toEqual([])
    recordSheetDuplicate(journal, 'u-copy', 'Data Copy', 'sheet-1')
    expect(toSaveSheetOps(journal)).toEqual([
      { kind: 'duplicate-sheet', sheetId: 'u-copy', name: 'Data Copy', sourceSheetId: 'sheet-1' },
    ])
    expect(toSaveEdits(journal).filter((edit) => edit.sheetId === 'u-copy')).toHaveLength(1)
  })
})

describe('filter dirty tracking', () => {
  it('counts one pending edit per dirty sheet and skips removed sheets', () => {
    const journal = createEditJournal()
    recordFilterChange(journal, 'sheet-1')
    recordFilterChange(journal, 'sheet-1')
    recordFilterChange(journal, 'sheet-2')
    expect(journalSize(journal)).toBe(2)
    recordSheetRemove(journal, 'sheet-2')
    expect(journalSize(journal)).toBe(2) // sheet-2 filter suppressed, its removal op counted instead
    expect(isSheetRemoved(journal, 'sheet-2')).toBe(true)
  })
})

describe('sheet visibility and order journal', () => {
  it('drops a hide toggled back to its original state and counts ops', () => {
    const journal = createEditJournal()
    recordSheetHidden(journal, 'sheet-1', true, false)
    expect(journalSize(journal)).toBe(1)
    recordSheetHidden(journal, 'sheet-1', false, false)
    expect(journalSize(journal)).toBe(0)
    recordSheetOrderChange(journal)
    expect(journalSize(journal)).toBe(1)
    recordSheetHidden(journal, 'sheet-2', true, false)
    expect(toSaveSheetOps(journal)).toEqual([
      { kind: 'set-sheet-hidden', sheetId: 'sheet-2', hidden: true },
      { kind: 'reorder-sheets' },
    ])
  })
})

describe('rich-text run capture', () => {
  it('captures per-run styling from a rich document mutation', () => {
    const journal = createEditJournal()
    const recorded = recordSetRangeValues(journal, 'sheet-1', {
      0: {
        0: {
          p: {
            body: {
              dataStream: 'Hello World\r\n',
              textRuns: [
                { st: 0, ed: 5, ts: { bl: 1, cl: { rgb: '#FF0000' } } },
                { st: 6, ed: 11, ts: { it: 1, fs: 14 } },
              ],
            },
          },
        },
      },
    })
    expect(recorded).toHaveLength(1)
    expect(recorded[0]?.value).toBe('Hello World')
    expect(recorded[0]?.rich).toEqual([
      {
        text: 'Hello',
        bold: true,
        italic: false,
        underline: false,
        strikethrough: false,
        color: '#FF0000',
      },
      { text: ' ', bold: false, italic: false, underline: false, strikethrough: false },
      {
        text: 'World',
        bold: false,
        italic: true,
        underline: false,
        strikethrough: false,
        size: 14,
      },
    ])
    const edits = toSaveEdits(journal)
    expect(edits[0]?.rich).toHaveLength(3)
  })

  it('drops runs when the cell is overwritten with a plain value', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', {
      0: {
        0: { p: { body: { dataStream: 'Rich\r\n', textRuns: [{ st: 0, ed: 4, ts: { bl: 1 } }] } } },
      },
    })
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { v: 'plain' } } })
    const edits = toSaveEdits(journal)
    expect(edits[0]?.value).toBe('plain')
    expect(edits[0]?.rich).toBeUndefined()
  })

  it('treats an unstyled rich document as plain text', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', {
      0: { 0: { p: { body: { dataStream: 'Plain\r\n', textRuns: [{ st: 0, ed: 5, ts: {} }] } } } },
    })
    expect(toSaveEdits(journal)[0]?.rich).toBeUndefined()
  })
})

describe('recordChartEdit', () => {
  it('merges series entries by index across successive edits', () => {
    const journal = createEditJournal()
    recordChartEdit(journal, 'xl/charts/chart1.xml', {
      series: [{ index: 0, values: [1, 2], valuesRef: "'Data'!$B$2:$B$3" }],
    })
    recordChartEdit(journal, 'xl/charts/chart1.xml', {
      title: 'Revenue',
      series: [
        { index: 0, categories: ['a', 'b'], categoriesRef: "'Data'!$A$2:$A$3" },
        { index: 1, name: 'East' },
      ],
    })
    const saved = toSaveChartEdits(journal)
    expect(saved).toHaveLength(1)
    expect(saved[0]?.title).toBe('Revenue')
    expect(saved[0]?.series).toEqual([
      {
        index: 0,
        values: [1, 2],
        valuesRef: "'Data'!$B$2:$B$3",
        categories: ['a', 'b'],
        categoriesRef: "'Data'!$A$2:$A$3",
      },
      { index: 1, name: 'East' },
    ])
  })

  it('later values for the same series win, refs and caches staying paired', () => {
    const journal = createEditJournal()
    recordChartEdit(journal, 'xl/charts/chart1.xml', {
      series: [{ index: 0, values: [1], valuesRef: "'Data'!$B$2:$B$2" }],
    })
    recordChartEdit(journal, 'xl/charts/chart1.xml', {
      series: [{ index: 0, values: [4, 5, 6], valuesRef: "'Data'!$C$2:$C$4" }],
    })
    const saved = toSaveChartEdits(journal)
    expect(saved[0]?.series).toEqual([
      { index: 0, values: [4, 5, 6], valuesRef: "'Data'!$C$2:$C$4" },
    ])
  })
})

describe('recordStructuralOp visual shifting', () => {
  const anchor = {
    fromRow: 10,
    fromColumn: 2,
    fromRowOffset: 5,
    fromColumnOffset: 5,
    toRow: 20,
    toColumn: 8,
    toRowOffset: 5,
    toColumnOffset: 5,
  }
  const chartVisual = {
    id: 'added-1',
    sheetId: 'sheet-1',
    kind: 'chart' as const,
    anchor,
    chart: {
      chartTypes: ['barChart'],
      title: 'T',
      series: [
        {
          name: 'S',
          categories: ['a', 'b'],
          values: [1, 2],
          valuesRef: "'Data'!$B$3:$B$4",
          categoriesRef: "'Data'!$A$3:$A$4",
        },
      ],
    },
  }

  it('shifts session-chart anchors and refs on a row insert', () => {
    const journal = createEditJournal()
    recordVisualAdd(journal, structuredClone(chartVisual))
    recordStructuralOp(journal, 'sheet-1', { kind: 'insert-rows', index: 1, count: 3 }, 'Data')
    const visual = journal.visualAdds[0]
    expect(visual?.anchor.fromRow).toBe(13)
    expect(visual?.anchor.toRow).toBe(23)
    expect(visual?.chart?.series[0]?.valuesRef).toBe("'Data'!$B$6:$B$7")
    expect(visual?.chart?.series[0]?.categoriesRef).toBe("'Data'!$A$6:$A$7")
  })

  it('clamps anchors into a removed span and drops destroyed refs', () => {
    const journal = createEditJournal()
    recordVisualAdd(journal, structuredClone(chartVisual))
    recordStructuralOp(journal, 'sheet-1', { kind: 'remove-rows', index: 2, count: 1 }, 'Data')
    let visual = journal.visualAdds[0]
    expect(visual?.anchor.fromRow).toBe(9)
    // Deleting a row inside the range shrinks it.
    expect(visual?.chart?.series[0]?.valuesRef).toBe("'Data'!$B$3")
    // Removing the whole referenced range drops the ref (literal fallback).
    recordStructuralOp(journal, 'sheet-1', { kind: 'remove-rows', index: 1, count: 2 }, 'Data')
    visual = journal.visualAdds[0]
    expect(visual?.chart?.series[0]?.valuesRef).toBeUndefined()
    expect(visual?.chart?.series[0]?.values).toEqual([1, 2])
  })

  it('leaves other sheets alone and shifts pending chart-edit refs', () => {
    const journal = createEditJournal()
    recordVisualAdd(journal, structuredClone({ ...chartVisual, sheetId: 'sheet-2' }))
    recordChartEdit(journal, 'xl/charts/chart1.xml', {
      series: [{ index: 0, values: [1], valuesRef: "'Data'!$B$2:$B$9" }],
    })
    recordStructuralOp(journal, 'sheet-1', { kind: 'insert-cols', index: 0, count: 2 }, 'Data')
    // Anchor untouched (other sheet), but the by-name ref still shifts.
    expect(journal.visualAdds[0]?.anchor.fromColumn).toBe(2)
    expect(journal.visualAdds[0]?.chart?.series[0]?.valuesRef).toBe("'Data'!$D$3:$D$4")
    expect(toSaveChartEdits(journal)[0]?.series?.[0]?.valuesRef).toBe("'Data'!$D$2:$D$9")
  })
})

describe('toRecalcUserInput', () => {
  it('maps journal entries to recalc user input', () => {
    const base = { row: 0, column: 0, hasValue: true }
    expect(toRecalcUserInput({ ...base, value: null, formula: '=SUM(A1:A2)' })).toBe('=SUM(A1:A2)')
    expect(toRecalcUserInput({ ...base, value: 42 })).toBe('42')
    expect(toRecalcUserInput({ ...base, value: true })).toBe('TRUE')
    expect(toRecalcUserInput({ ...base, value: false })).toBe('FALSE')
    expect(toRecalcUserInput({ ...base, value: null })).toBe('')
    expect(toRecalcUserInput({ ...base, value: 'plain' })).toBe('plain')
  })

  it('escapes text that the engine would reinterpret as a formula', () => {
    const base = { row: 0, column: 0, hasValue: true }
    expect(toRecalcUserInput({ ...base, value: '=not a formula' })).toBe("'=not a formula")
    expect(toRecalcUserInput({ ...base, value: "'quoted" })).toBe("''quoted")
  })
})

describe('recordStructuralOp move-rows', () => {
  it('remaps journaled cells through the swap and cancels the inverse move', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', {
      1: { 0: { v: 'moved' } },
      3: { 0: { v: 'displaced' } },
      7: { 0: { v: 'outside' } },
    })
    recordStructuralOp(journal, 'sheet-1', { kind: 'move-rows', index: 1, count: 1, before: 5 })
    const cells = journal.cells.get('sheet-1')
    expect(cells?.get('4:0')?.value).toBe('moved')
    expect(cells?.get('2:0')?.value).toBe('displaced')
    expect(cells?.get('7:0')?.value).toBe('outside')

    // Undo arrives as the exact inverse move and cancels the pair.
    recordStructuralOp(journal, 'sheet-1', { kind: 'move-rows', index: 4, count: 1, before: 1 })
    expect(journal.structuralOps.get('sheet-1')).toBeUndefined()
    const restored = journal.cells.get('sheet-1')
    expect(restored?.get('1:0')?.value).toBe('moved')
    expect(restored?.get('3:0')?.value).toBe('displaced')
  })
})

describe('workbook protection / theme / protected-range journaling', () => {
  it('workbook protection drops when toggled back to the original', () => {
    const journal = createEditJournal()
    recordWorkbookProtection(journal, true, false)
    expect(journal.workbookProtection.desired).toBe(true)
    expect(journalSize(journal)).toBe(1)
    recordWorkbookProtection(journal, false, false)
    expect(journal.workbookProtection.desired).toBeNull()
    expect(journalSize(journal)).toBe(0)
  })

  it('theme colors and fonts count once each and overwrite on re-pick', () => {
    const journal = createEditJournal()
    recordThemeColors(
      journal,
      'Indigo',
      Array.from({ length: 12 }, () => '#111111'),
    )
    recordThemeFonts(journal, 'Georgia', 'Georgia', 'Georgia')
    expect(journalSize(journal)).toBe(2)
    recordThemeColors(
      journal,
      'Forest',
      Array.from({ length: 12 }, () => '#222222'),
    )
    expect(journal.theme.colors?.name).toBe('Forest')
    expect(journalSize(journal)).toBe(2)
  })

  it('protected-range dirt skips removed sheets in the size count', () => {
    const journal = createEditJournal()
    recordProtectedRangesChange(journal, 'sheet-1')
    recordProtectedRangesChange(journal, 'sheet-2')
    expect(journalSize(journal)).toBe(2)
    recordSheetRemove(journal, 'sheet-2')
    expect(journalSize(journal)).toBe(2) // sheet-2's dirt swaps for the removal op
  })
})

describe('recordStructuralOp page-break remapping', () => {
  it('shifts journaled breaks through inserts and drops deleted ones', () => {
    const journal = createEditJournal()
    recordPageSetup(journal, 'sheet-1', { rowBreaks: [3, 8], colBreaks: [2] })
    recordStructuralOp(journal, 'sheet-1', { kind: 'insert-rows', index: 5, count: 2 })
    expect(journal.pageSetup.get('sheet-1')?.rowBreaks).toEqual([3, 10])
    expect(journal.pageSetup.get('sheet-1')?.colBreaks).toEqual([2])
    recordStructuralOp(journal, 'sheet-1', { kind: 'remove-rows', index: 2, count: 3 })
    expect(journal.pageSetup.get('sheet-1')?.rowBreaks).toEqual([7])
  })

  it('undo (the inverse op) restores the shifted break set', () => {
    const journal = createEditJournal()
    recordPageSetup(journal, 'sheet-1', { rowBreaks: [6] })
    recordStructuralOp(journal, 'sheet-1', { kind: 'insert-rows', index: 0, count: 4 })
    expect(journal.pageSetup.get('sheet-1')?.rowBreaks).toEqual([10])
    recordStructuralOp(journal, 'sheet-1', { kind: 'remove-rows', index: 0, count: 4 })
    expect(journal.structuralOps.get('sheet-1')).toBeUndefined()
    expect(journal.pageSetup.get('sheet-1')?.rowBreaks).toEqual([6])
  })
})
