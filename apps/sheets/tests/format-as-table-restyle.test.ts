import JSZip from 'jszip'
import { describe, expect, it, vi } from 'vitest'
import { t } from '../src/renderer/i18n/locale'
import { createEditJournal } from '../src/renderer/edit-journal'
import { handleRibbonCommand, type RibbonCommandContext } from '../src/renderer/ribbon-actions'
import type { DataToolsContext } from '../src/renderer/data-tools-actions'
import type { TableDesignContext } from '../src/renderer/table-design-actions'
import type { LazyWorkbookState, UniverRuntime } from '../src/renderer/univer-state'
import { buildEditFixture, buildKitchenSinkFixture } from './fixture-builder'
import {
  createBufferEntrySource,
  planCellEditsToXlsx,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'

vi.mock('../src/renderer/univer-sync', async (original) => ({
  ...(await original<typeof import('../src/renderer/univer-sync')>()),
  loadVisibleRange: vi.fn(() => Promise.resolve()),
}))
const flush = () => new Promise((resolve) => setTimeout(resolve, 0))

function harness(fileTable = false, named = true) {
  const messages: string[] = []
  const themes: unknown[] = []
  const tables = new Map<string, string>()
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
    addTable: vi.fn(async (name: string, _area: unknown, id: string) => {
      tables.set(name, id)
      return true
    }),
    addTableTheme: vi.fn(async (_id: string, theme: unknown) => {
      themes.push(theme)
      return true
    }),
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
    getTableInfoByName: (name: string) => (tables.has(name) ? { id: tables.get(name) } : undefined),
  }
  const state = {
    file: {
      sheets: [
        {
          id: 's1',
          name: 'Data',
          tables: fileTable
            ? [
                {
                  name: 'Table1',
                  range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
                  columns: ['A', 'hdr'],
                  headerRowCount: 1,
                  showRowStripes: true,
                  styleName: 'TableStyleMedium2',
                },
              ]
            : [],
        },
      ],
      definedNames: named ? [{ name: 'SalesData', ref: "'Data'!$A$1:$B$4" }] : [],
    },
    editJournal: createEditJournal(),
    loadedRanges: new Map(),
    frozenStripKeys: new Set(),
  } as unknown as LazyWorkbookState
  const refs = {
    univerRef: {
      current: { univerAPI: { getActiveWorkbook: () => workbook } } as unknown as UniverRuntime,
    },
    lazyWorkbookRef: { current: state },
  }
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
    messages,
    worksheet,
    design,
    themes,
    select: (row: number, column: number, height = 1, width = 1) => {
      selection = { row, column, height, width }
    },
    pick: (style: string) => handleRibbonCommand(ctx, `format-as-table:${style}`),
  }
}

async function saveTables(h: ReturnType<typeof harness>, fileTable: boolean) {
  const buffer = await (fileTable ? buildKitchenSinkFixture() : buildEditFixture())
  const source = await createBufferEntrySource(buffer)
  const args: unknown[] = [source, [], []]
  while (args.length < 27)
    args.push(args.length === 10 || args.length === 21 || args.length === 22 ? null : [])
  args[4] = undefined
  args[14] = h.state.editJournal.tableAdds.map(({ sheetId: _id, ...add }) => ({
    sheetName: 'Data',
    ...add,
  }))
  args[26] = [...h.state.editJournal.tableEdits.values()].map(({ sheetId: _id, ...edit }) => ({
    sheetName: 'Data',
    ...edit,
  }))
  const plan = await planCellEditsToXlsx(...(args as Parameters<typeof planCellEditsToXlsx>))
  const zip = await JSZip.loadAsync(buffer)
  for (const path of plan.removedEntries) zip.remove(path)
  for (const [path, xml] of [...plan.added, ...plan.replaced]) zip.file(path, xml)
  for (const [path, bytes] of plan.addedBinary) zip.file(path, bytes)
  const reopened = await JSZip.loadAsync(await zip.generateAsync({ type: 'uint8array' }))
  const tableParts = Object.keys(reopened.files).filter((p) =>
    /^xl\/tables\/table\d+\.xml$/.test(p),
  )
  expect(tableParts).toHaveLength(1)
  expect(await reopened.file(tableParts[0]!)!.async('string')).toContain('name="TableStyleDark2"')
  return plan
}

describe('Format as Table on an existing table', () => {
  it.each([false, true])(
    'restyles the session table once with named selection=%s and saves the new style',
    async (named) => {
      const h = harness(false, named)
      const names = structuredClone(h.state.file.definedNames)
      h.pick('TableStyleMedium2')
      await flush()
      expect(h.state.editJournal.tableAdds).toHaveLength(1)
      const created = structuredClone(h.state.editJournal.tableAdds[0])
      h.pick('TableStyleDark2')
      await flush()
      expect(h.state.editJournal.tableAdds).toHaveLength(1)
      expect(h.state.editJournal.tableAdds[0]).toEqual({ ...created, style: 'TableStyleDark2' })
      expect(h.worksheet.addTable).toHaveBeenCalledTimes(1)
      expect(h.worksheet.addTableTheme).toHaveBeenCalledTimes(2)
      expect(h.design.refreshSelection).toHaveBeenCalledTimes(1)
      expect(h.messages).not.toContain(t('appTableOverlapsSession', { name: 'Table1' }))
      expect(h.state.file.definedNames).toEqual(names)
      const plan = await saveTables(h, false)
      const xml = [...plan.added.entries()].find(([path]) =>
        /^xl\/tables\/table\d+\.xml$/.test(path),
      )?.[1]
      expect(xml).toContain('name="TableStyleDark2"')
    },
  )

  it('restyles from a single cell without trying to infer and create another region', async () => {
    const h = harness()
    h.pick('TableStyleMedium2')
    await flush()
    h.select(2, 1)
    h.pick('TableStyleLight9')
    await flush()
    expect(h.state.editJournal.tableAdds).toHaveLength(1)
    expect(h.state.editJournal.tableAdds[0]!.style).toBe('TableStyleLight9')
  })

  it('edits a loaded table and saves a table edit instead of adding a duplicate', async () => {
    const h = harness(true)
    h.select(0, 0, 2, 2)
    h.pick('TableStyleDark2')
    await flush()
    expect(h.state.editJournal.tableAdds).toEqual([])
    expect([...h.state.editJournal.tableEdits.values()]).toEqual([
      { sheetId: 's1', tableName: 'Table1', style: 'TableStyleDark2' },
    ])
    expect(h.state.file.sheets[0]!.tables![0]!.styleName).toBe('TableStyleDark2')
    const plan = await saveTables(h, true)
    expect(plan.replaced.get('xl/tables/table1.xml')).toContain('name="TableStyleDark2"')
  })

  it('still rejects a new selection whose anchor is outside but overlaps an existing table', async () => {
    const h = harness()
    h.select(1, 1, 4, 2)
    h.pick('TableStyleMedium2')
    await flush()
    h.select(0, 0, 3, 2)
    h.pick('TableStyleDark2')
    await flush()
    expect(h.state.editJournal.tableAdds).toHaveLength(1)
    expect(h.state.editJournal.tableAdds[0]!.style).toBe('TableStyleMedium2')
    expect(h.messages).toContain(t('appTableOverlapsSession', { name: 'Table1' }))
  })
})
