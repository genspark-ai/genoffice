import { describe, expect, it } from 'vitest'

import {
  createBufferEntrySource,
  planCellEditsToXlsx,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { SheetTableEdit } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { buildKitchenSinkFixture } from './fixture-builder'

// The kitchen-sink fixture carries Table1 (A1:B2, columns A / hdr) on sheet Data.
async function planWith(
  edits: SheetTableEdit[],
  structuralOps: Parameters<typeof planCellEditsToXlsx>[2] = [],
) {
  const source = await createBufferEntrySource(await buildKitchenSinkFixture())
  const args: unknown[] = [source, [], structuralOps]
  while (args.length < 27) args.push(args.length === 10 ? null : [])
  args[4] = undefined
  args[21] = null
  args[22] = null
  args[26] = edits
  return planCellEditsToXlsx(...(args as Parameters<typeof planCellEditsToXlsx>))
}

function edit(overrides: Partial<SheetTableEdit> = {}): SheetTableEdit {
  return { sheetName: 'Data', tableName: 'Table1', ...overrides }
}

describe('table edits', () => {
  it('renames the table part by displayName', async () => {
    const plan = await planWith([edit({ name: 'Sales' })])
    const xml = plan.replaced.get('xl/tables/table1.xml')
    expect(xml).toContain('name="Sales" displayName="Sales"')
    expect(xml).toContain('ref="A1:B2"')
  })

  it('writes the style and option flags into tableStyleInfo', async () => {
    const plan = await planWith([
      edit({
        style: 'TableStyleLight9',
        bandedRows: false,
        firstColumn: true,
        bandedColumns: true,
      }),
    ])
    const xml = plan.replaced.get('xl/tables/table1.xml')
    expect(xml).toContain(
      '<tableStyleInfo showFirstColumn="1" showLastColumn="0" showRowStripes="0" showColumnStripes="1" name="TableStyleLight9"/>',
    )
  })

  it('adds a totals row: count, shown flag and a filter that stops above it', async () => {
    const plan = await planWith([
      edit({
        totalsRow: true,
        filterButton: true,
        area: { startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 },
      }),
    ])
    const xml = plan.replaced.get('xl/tables/table1.xml')
    expect(xml).toContain('ref="A1:B3"')
    expect(xml).toContain('totalsRowCount="1"')
    expect(xml).toContain('totalsRowShown="1"')
    expect(xml).toContain('<autoFilter ref="A1:B2"/>')
  })

  it('turning totals off clears totalsRowShown too', async () => {
    const plan = await planWith([edit({ totalsRow: false })])
    const xml = plan.replaced.get('xl/tables/table1.xml')
    expect(xml).not.toContain('totalsRowCount')
    expect(xml).toContain('totalsRowShown="0"')
  })

  it('drops the header row and its filter', async () => {
    const plan = await planWith([
      edit({ headerRow: false, area: { startRow: 1, startColumn: 0, endRow: 2, endColumn: 1 } }),
    ])
    const xml = plan.replaced.get('xl/tables/table1.xml')
    expect(xml).toContain('headerRowCount="0"')
    expect(xml).not.toContain('<autoFilter')
  })

  it('resizes with a new column list, keeping existing column ids', async () => {
    const plan = await planWith([
      edit({
        area: { startRow: 0, startColumn: 0, endRow: 4, endColumn: 2 },
        columnNames: ['A', 'hdr', 'Notes'],
      }),
    ])
    const xml = plan.replaced.get('xl/tables/table1.xml')
    expect(xml).toContain('ref="A1:C5"')
    expect(xml).toContain(
      '<tableColumns count="3"><tableColumn id="1" name="A"/><tableColumn id="2" name="hdr"/><tableColumn id="3" name="Notes"/></tableColumns>',
    )
  })

  it('keeps filter criteria on edits that leave the columns alone', async () => {
    const source = await createBufferEntrySource(await buildKitchenSinkFixture())
    const xmlWithCriteria = (await source.readText('xl/tables/table1.xml')).replace(
      '<tableColumns',
      '<autoFilter ref="A1:B2"><filterColumn colId="0"><filters><filter val="1"/></filters></filterColumn></autoFilter><tableColumns',
    )
    const patched = {
      ...source,
      readText: (path: string) =>
        path === 'xl/tables/table1.xml' ? Promise.resolve(xmlWithCriteria) : source.readText(path),
    }
    const args: unknown[] = [patched, [], []]
    while (args.length < 27) args.push(args.length === 10 ? null : [])
    args[4] = undefined
    args[21] = null
    args[22] = null
    args[26] = [
      edit({ name: 'Sales', area: { startRow: 0, startColumn: 0, endRow: 4, endColumn: 1 } }),
    ]
    const plan = await planCellEditsToXlsx(...(args as Parameters<typeof planCellEditsToXlsx>))
    const xml = plan.replaced.get('xl/tables/table1.xml')
    expect(xml).toContain('<autoFilter ref="A1:B5"><filterColumn colId="0">')
  })

  it('keeps column ids with their sheet column when the left edge moves', async () => {
    const plan = await planWith([
      edit({
        area: { startRow: 0, startColumn: 1, endRow: 2, endColumn: 2 },
        columnNames: ['hdr', 'Notes'],
      }),
    ])
    const xml = plan.replaced.get('xl/tables/table1.xml')
    expect(xml).toContain(
      '<tableColumns count="2"><tableColumn id="2" name="hdr"/><tableColumn id="3" name="Notes"/></tableColumns>',
    )
    expect(xml).toContain('ref="B1:C3"')
  })

  it('removes the part, relationship, tablePart and content type for Convert to Range', async () => {
    const plan = await planWith([edit({ remove: true })])
    expect(plan.removedEntries).toContain('xl/tables/table1.xml')
    expect(plan.replaced.get('xl/worksheets/sheet1.xml')).not.toContain('<tableParts')
    expect(plan.replaced.get('xl/worksheets/_rels/sheet1.xml.rels')).not.toContain(
      'tables/table1.xml',
    )
    expect(plan.replaced.get('[Content_Types].xml')).not.toContain('/xl/tables/table1.xml')
  })

  it('applies edits before additions so a converted table frees its cells and name', async () => {
    const source = await createBufferEntrySource(await buildKitchenSinkFixture())
    const args: unknown[] = [source, [], []]
    while (args.length < 27) args.push(args.length === 10 ? null : [])
    args[4] = undefined
    args[21] = null
    args[22] = null
    args[14] = [
      {
        sheetName: 'Data',
        area: { startRow: 0, startColumn: 3, endRow: 1, endColumn: 4 },
        name: 'Table1',
        columnNames: ['A', 'hdr'],
        bandedRows: true,
      },
    ]
    args[26] = [edit({ remove: true })]
    const plan = await planCellEditsToXlsx(...(args as Parameters<typeof planCellEditsToXlsx>))
    expect(plan.removedEntries).toContain('xl/tables/table1.xml')
    expect([...plan.added.keys()].some((path) => path.startsWith('xl/tables/'))).toBe(true)
  })

  it('fails closed on unknown tables, bad names and width mismatches', async () => {
    await expect(planWith([edit({ tableName: 'Nope', name: 'X' })])).rejects.toThrow(/not found/)
    await expect(planWith([edit({ name: 'A1' })])).rejects.toThrow(/not a valid table name/)
    await expect(
      planWith([edit({ area: { startRow: 0, startColumn: 0, endRow: 2, endColumn: 3 } })]),
    ).rejects.toThrow(/spans 4 columns/)
  })

  it('fails closed when a resize rides with row shifts on the same sheet', async () => {
    await expect(
      planWith(
        [edit({ area: { startRow: 0, startColumn: 0, endRow: 4, endColumn: 1 } })],
        [{ sheetName: 'Data', ops: [{ kind: 'insert-rows', index: 0, count: 1 }] }],
      ),
    ).rejects.toThrow(/save the table first/)
  })
})
