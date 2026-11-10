import { describe, expect, it } from 'vitest'

import {
  handleCreateNamesFromSelection,
  type DataToolsContext,
} from '../src/renderer/data-tools-actions'

interface Selection {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}

function makeHarness(options: {
  cells: string[][]
  selection: Selection
  existingNames?: string[]
  /// Simulates lazy loading: Univer's raw cell data reports no content even
  /// though the file (and getDisplayValue) has it.
  univerSeesNoContent?: boolean
}) {
  const inserted: { name: string; ref: string }[] = []
  const messages: string[] = []
  let reads = 0
  const lastRow = options.univerSeesNoContent ? -1 : options.cells.length - 1
  const lastColumn = options.univerSeesNoContent
    ? -1
    : Math.max(...options.cells.map((row) => row.length)) - 1
  const worksheet = {
    getSheetId: () => 'sheet1',
    getSheetName: () => 'Data',
    getLastRow: () => lastRow,
    getLastColumn: () => lastColumn,
    getRange: (row: number, column: number) => ({
      getDisplayValue: () => {
        reads += 1
        return options.cells[row]?.[column] ?? ''
      },
    }),
  }
  const workbook = {
    getActiveSheet: () => worksheet,
    getActiveRange: () => ({ getRange: () => options.selection }),
    getDefinedNames: () => (options.existingNames ?? []).map((name) => ({ getName: () => name })),
    insertDefinedName: (name: string, ref: string) => {
      inserted.push({ name, ref })
    },
  }
  const file = {
    sheets: [
      {
        id: 'sheet1',
        rowCount: options.cells.length,
        columnCount: Math.max(...options.cells.map((row) => row.length)),
      },
    ],
  }
  const ctx = {
    univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
    lazyWorkbookRef: { current: { file } },
    setMessage: (message: string) => {
      messages.push(message)
    },
  } as unknown as DataToolsContext
  return { ctx, inserted, messages, readCount: () => reads }
}

describe('handleCreateNamesFromSelection', () => {
  it('creates one name per column from top-row labels, Unicode included', () => {
    const { ctx, inserted } = makeHarness({
      cells: [
        ['销售额', 'Region'],
        ['1', 'a'],
        ['2', 'b'],
      ],
      selection: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
    })
    handleCreateNamesFromSelection(ctx, 'top')
    expect(inserted).toEqual([
      { name: '销售额', ref: "'Data'!$A$2:$A$3" },
      { name: 'Region', ref: "'Data'!$B$2:$B$3" },
    ])
  })

  it('skips names that already exist, case-insensitively', () => {
    // insertDefinedName accepts duplicates silently (internal-id keyed) and
    // the duplicate then fails the save — the skip must happen here.
    const { ctx, inserted } = makeHarness({
      cells: [
        ['Total', 'Fresh'],
        ['1', '2'],
      ],
      selection: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
      existingNames: ['tOtAl'],
    })
    handleCreateNamesFromSelection(ctx, 'top')
    expect(inserted.map((entry) => entry.name)).toEqual(['Fresh'])
  })

  it('skips duplicate labels within the batch', () => {
    const { ctx, inserted } = makeHarness({
      cells: [
        ['Total', 'Total'],
        ['1', '2'],
      ],
      selection: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
    })
    handleCreateNamesFromSelection(ctx, 'top')
    expect(inserted).toHaveLength(1)
  })

  it('walks labels only over the data extent on whole-column selections', () => {
    const { ctx, inserted, readCount } = makeHarness({
      cells: [
        ['Sales', 'Year'],
        ['1', '2020'],
        ['2', '2021'],
      ],
      selection: { startRow: 0, endRow: 1_048_575, startColumn: 0, endColumn: 16_383 },
    })
    handleCreateNamesFromSelection(ctx, 'top')
    // Refs keep the full selection like Excel — future data lands in the
    // named range; only the label reads stop at the data extent.
    expect(inserted).toEqual([
      { name: 'Sales', ref: "'Data'!$A$2:$A$1048576" },
      { name: 'Year', ref: "'Data'!$B$2:$B$1048576" },
    ])
    expect(readCount()).toBe(2)
  })

  it('creates names over empty data rows below the labels', () => {
    // Labels-only selection (data typed later): Excel still creates the
    // names, spanning the selected empty rows.
    const { ctx, inserted } = makeHarness({
      cells: [['Old', '10']],
      selection: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
    })
    handleCreateNamesFromSelection(ctx, 'top')
    expect(inserted).toEqual([
      { name: 'Old', ref: "'Data'!$A$2:$A$3" },
      { name: '_10', ref: "'Data'!$B$2:$B$3" },
    ])
  })

  it('uses the file extent when Univer has not materialized the cells yet', () => {
    // Lazily opened workbooks serve file cells through interceptors, so the
    // raw-content last row reads -1; the clamp must fall back to the file's
    // used range instead of refusing with "needs data".
    const { ctx, inserted } = makeHarness({
      cells: [
        ['Sales', 'Year'],
        ['1', '2020'],
        ['2', '2021'],
      ],
      selection: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
      univerSeesNoContent: true,
    })
    handleCreateNamesFromSelection(ctx, 'top')
    expect(inserted.map((entry) => entry.name)).toEqual(['Sales', 'Year'])
  })

  it('walks labels only over the data extent in left mode', () => {
    const { ctx, inserted, readCount } = makeHarness({
      cells: [
        ['North', '1', '2'],
        ['South', '3', '4'],
      ],
      selection: { startRow: 0, endRow: 1_048_575, startColumn: 0, endColumn: 16_383 },
    })
    handleCreateNamesFromSelection(ctx, 'left')
    expect(inserted).toEqual([
      { name: 'North', ref: "'Data'!$B$1:$XFD$1" },
      { name: 'South', ref: "'Data'!$B$2:$XFD$2" },
    ])
    expect(readCount()).toBe(2)
  })
})
