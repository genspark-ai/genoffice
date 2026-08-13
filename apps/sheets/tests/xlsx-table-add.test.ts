import { describe, expect, it } from 'vitest'

import { createBufferEntrySource, planCellEditsToXlsx } from '../src/gateway/xlsx-gateway'
import type { SheetTableAddition } from '../src/gateway/xlsx-gateway'
import { buildEditFixture, buildStructureFixture } from './fixture-builder'

function tableAddition(overrides: Partial<SheetTableAddition> = {}): SheetTableAddition {
  return {
    sheetName: 'Data',
    area: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 1 },
    name: 'Users',
    columnNames: ['Country', 'Users'],
    bandedRows: true,
    ...overrides,
  }
}

async function planWith(
  tables: SheetTableAddition[],
  buffer?: Buffer,
  structuralOps: Parameters<typeof planCellEditsToXlsx>[2] = [],
) {
  const source = await createBufferEntrySource(buffer ?? (await buildEditFixture()))
  return planCellEditsToXlsx(
    source,
    [],
    structuralOps,
    [],
    undefined,
    [],
    [],
    [],
    [],
    [],
    null,
    [],
    [],
    [],
    tables,
  )
}

describe('table additions', () => {
  it('creates the table part and registers every hook', async () => {
    const plan = await planWith([tableAddition()])

    const tableXml = plan.added.get('xl/tables/table1.xml')
    expect(tableXml).toBeDefined()
    expect(tableXml).toContain('id="1" name="Users" displayName="Users" ref="A1:B4"')
    expect(tableXml).toContain('<autoFilter ref="A1:B4"/>')
    expect(tableXml).toContain('<tableColumns count="2">')
    expect(tableXml).toContain('<tableColumn id="1" name="Country"/>')
    expect(tableXml).toContain('<tableColumn id="2" name="Users"/>')
    expect(tableXml).toContain('name="TableStyleMedium2"')
    expect(tableXml).toContain('showRowStripes="1"')

    const worksheet = plan.replaced.get('xl/worksheets/sheet1.xml')
    expect(worksheet).toContain('<tableParts count="1"><tablePart r:id="rId1"/></tableParts>')
    // The fixture worksheet has no xmlns:r — the table hookup must add it.
    expect(worksheet).toContain(
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
    )

    const sheetRels = plan.added.get('xl/worksheets/_rels/sheet1.xml.rels')
    expect(sheetRels).toContain('relationships/table')
    expect(sheetRels).toContain('Target="../tables/table1.xml"')

    const contentTypes = plan.replaced.get('[Content_Types].xml')
    expect(contentTypes).toContain('PartName="/xl/tables/table1.xml"')
    expect(contentTypes).toContain('spreadsheetml.table+xml')
  })

  it('honors style and banding options and escapes names', async () => {
    const plan = await planWith([
      tableAddition({
        name: 'Sales_Q3',
        style: 'TableStyleLight9',
        bandedRows: false,
        columnNames: ['A<B', 'C"D'],
      }),
    ])
    const tableXml = plan.added.get('xl/tables/table1.xml')
    expect(tableXml).toContain('name="Sales_Q3" displayName="Sales_Q3"')
    expect(tableXml).toContain('name="TableStyleLight9"')
    expect(tableXml).toContain('showRowStripes="0"')
    expect(tableXml).toContain('name="A&lt;B"')
    expect(tableXml).toContain('name="C&quot;D"')
  })

  it('adds two disjoint tables with distinct parts, ids, and rel ids', async () => {
    const plan = await planWith([
      tableAddition(),
      tableAddition({
        name: 'Costs',
        area: { startRow: 0, startColumn: 3, endRow: 3, endColumn: 4 },
        columnNames: ['Item', 'Cost'],
      }),
    ])
    expect(plan.added.get('xl/tables/table1.xml')).toContain('id="1" name="Users"')
    expect(plan.added.get('xl/tables/table2.xml')).toContain('id="2" name="Costs"')
    const worksheet = plan.replaced.get('xl/worksheets/sheet1.xml')
    expect(worksheet).toContain('<tableParts count="2">')
    expect(worksheet).toContain('<tablePart r:id="rId1"/>')
    expect(worksheet).toContain('<tablePart r:id="rId2"/>')
  })

  it('rejects overlapping tables in the same save', async () => {
    await expect(
      planWith([
        tableAddition(),
        tableAddition({
          name: 'Clash',
          area: { startRow: 2, startColumn: 1, endRow: 5, endColumn: 3 },
          columnNames: ['X', 'Y', 'Z'],
        }),
      ]),
    ).rejects.toThrow(/overlaps an existing table/)
  })

  it('rejects duplicate and defined-name-colliding table names', async () => {
    await expect(
      planWith([
        tableAddition(),
        tableAddition({
          name: 'users',
          area: { startRow: 0, startColumn: 3, endRow: 3, endColumn: 3 },
          columnNames: ['X'],
        }),
      ]),
    ).rejects.toThrow(/already taken/)
    // The edit fixture defines the workbook name "Total".
    await expect(planWith([tableAddition({ name: 'Total' })])).rejects.toThrow(
      /collides with a defined name/,
    )
  })

  it('rejects bad column names and header-only areas', async () => {
    await expect(planWith([tableAddition({ columnNames: ['Country', ' '] })])).rejects.toThrow(
      /blank column name/,
    )
    await expect(planWith([tableAddition({ columnNames: ['Same', 'same'] })])).rejects.toThrow(
      /duplicate column name/,
    )
    await expect(planWith([tableAddition({ columnNames: ['Only'] })])).rejects.toThrow(
      /spans 2 columns/,
    )
    await expect(
      planWith([
        tableAddition({
          area: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 },
        }),
      ]),
    ).rejects.toThrow(/at least one data row/)
  })

  it('rejects a table over merged cells', async () => {
    await expect(
      planWith(
        [
          tableAddition({
            area: { startRow: 3, startColumn: 0, endRow: 6, endColumn: 1 },
            columnNames: ['X', 'Y'],
          }),
        ],
        await buildStructureFixture(),
      ),
    ).rejects.toThrow(/overlaps merged cells/)
  })

  it('fails closed when the same save shifts rows on the table sheet', async () => {
    await expect(
      planWith([tableAddition()], undefined, [
        { sheetName: 'Data', ops: [{ kind: 'insert-rows', index: 0, count: 1 }] },
      ]),
    ).rejects.toThrow(/save the table first/)
  })
})
