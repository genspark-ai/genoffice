import { describe, expect, it, vi } from 'vitest'
import { t } from '../src/renderer/i18n/locale'
import { createEditJournal, recordTableEdit } from '../src/renderer/edit-journal'
import { handleRibbonCommand, type RibbonCommandContext } from '../src/renderer/ribbon-actions'
import { applyAiTableAdd } from '../src/renderer/workbook-ops'
import type { DataToolsContext } from '../src/renderer/data-tools-actions'
import type { TableDesignContext } from '../src/renderer/table-design-actions'
import type { LazyWorkbookState, UniverRuntime } from '../src/renderer/univer-state'

vi.mock('../src/renderer/univer-sync', async (original) => ({
  ...(await original<typeof import('../src/renderer/univer-sync')>()),
  loadVisibleRange: vi.fn(() => Promise.resolve()),
}))
const flush = () => new Promise((resolve) => setTimeout(resolve, 0))

/// Table1 from the file sits on rows 2-3; the data block spans rows 1-4.
function harness() {
  const messages: string[] = []
  let selection = { row: 0, column: 0, height: 4, width: 2 }
  const cells = [
    ['Item', 'Amount'],
    ['coffee', 1],
    ['tea', 2],
    ['water', 3],
  ]
  const worksheet = {
    getSheetId: () => 's1',
    getLastRow: () => cells.length - 1,
    getLastColumn: () => 1,
    getRange: (row: number, column: number, height = 1, width = 1) => ({
      getValue: () => cells[row]?.[column] ?? null,
      getDisplayValue: () => String(cells[row]?.[column] ?? ''),
      getValues: () =>
        Array.from({ length: height }, (_, r) =>
          Array.from({ length: width }, (_, c) => cells[row + r]?.[column + c] ?? null),
        ),
      setValue: vi.fn(),
    }),
    addTable: vi.fn(async () => true),
    addTableTheme: vi.fn(async () => true),
  }
  const workbook = {
    getActiveSheet: () => worksheet,
    getSheetBySheetId: () => worksheet,
    getActiveRange: () => ({
      getRow: () => selection.row,
      getColumn: () => selection.column,
      getHeight: () => selection.height,
      getWidth: () => selection.width,
    }),
    getTableInfoByName: () => undefined,
  }
  const state = {
    file: {
      sheets: [
        {
          id: 's1',
          name: 'Data',
          tables: [
            {
              name: 'Table1',
              range: { startRow: 1, endRow: 2, startColumn: 0, endColumn: 1 },
              columns: ['coffee', '1'],
              headerRowCount: 1,
              showRowStripes: true,
              styleName: 'TableStyleMedium2',
            },
          ],
        },
      ],
      definedNames: [],
    },
    editJournal: createEditJournal(),
    loadedRanges: new Map(),
    frozenStripKeys: new Set(),
  } as unknown as LazyWorkbookState
  const runtime = { univerAPI: { getActiveWorkbook: () => workbook } } as unknown as UniverRuntime
  const refs = { univerRef: { current: runtime }, lazyWorkbookRef: { current: state } }
  const common = {
    ...refs,
    setMessage: (m: string) => {
      messages.push(m)
    },
    setPendingEdits: vi.fn(),
    setAdvancedFilterColumns: vi.fn(),
  }
  const design = { ...common, refreshSelection: vi.fn() } as TableDesignContext
  const ctx = {
    ...common,
    dataToolsContext: () => common as DataToolsContext,
    tableDesignContext: () => design,
  } as unknown as RibbonCommandContext
  return {
    state,
    runtime,
    messages,
    worksheet,
    select: (row: number, column: number, height = 1, width = 1) => {
      selection = { row, column, height, width }
    },
    run: (command: string) => handleRibbonCommand(ctx, command),
  }
}

describe('creating a table over a table loaded from the file', () => {
  it('Home gallery rejects a selection that starts outside Table1 but covers it', async () => {
    const h = harness()
    h.run('format-as-table:TableStyleDark2')
    await flush()
    expect(h.state.editJournal.tableAdds).toEqual([])
    expect(h.worksheet.addTable).not.toHaveBeenCalled()
    expect(h.messages).toContain(t('appTableOverlapsFile', { name: 'Table1' }))
  })

  it('Insert → Table inside Table1 explains instead of stacking a second table', async () => {
    const h = harness()
    h.select(2, 1)
    h.run('format-as-table')
    await flush()
    expect(h.state.editJournal.tableAdds).toEqual([])
    expect(h.messages).toEqual([t('appTableAlreadyInTable', { name: 'Table1' })])
  })

  it('AI add_table over Table1 throws', () => {
    const h = harness()
    expect(() =>
      applyAiTableAdd(h.runtime, h.state, {
        op: 'add_table',
        sheetId: 's1',
        range: 'A1:B4',
        style: 'TableStyleMedium2',
      }),
    ).toThrow(t('appTableOverlapsFile', { name: 'Table1' }))
    expect(h.state.editJournal.tableAdds).toEqual([])
  })

  it('a file table removed this session no longer blocks the range', async () => {
    const h = harness()
    recordTableEdit(h.state.editJournal, { sheetId: 's1', tableName: 'Table1', remove: true })
    h.run('format-as-table:TableStyleDark2')
    await flush()
    expect(h.state.editJournal.tableAdds).toHaveLength(1)
    expect(h.messages).toContain(t('appTableCreated'))
  })
})
