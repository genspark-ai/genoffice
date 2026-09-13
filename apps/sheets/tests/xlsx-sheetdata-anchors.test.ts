/**
 * A save reported "Worksheet has no sheetData element." with no hint which
 * part or why. Two causes: the table header sync's patchCell only knew the
 * `</sheetData>` form (an added sheet starts as `<sheetData/>`), and an
 * empty/truncated worksheet part produced the same bare message.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { blankXlsxBuffer } from '@genoffice/xlsx-gateway/gateway/csv-import'
import {
  assembleWithJsZip,
  createBufferEntrySource,
  planCellEditsToXlsx,
  type CellEdit,
  type SheetStructuralOps,
  type SheetTableAddition,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { SheetEditPlan } from '@genoffice/xlsx-gateway/gateway/xlsx-sheets'

async function save(
  source: Buffer,
  edits: readonly CellEdit[],
  structural: readonly SheetStructuralOps[],
  sheetPlan: SheetEditPlan | undefined,
  tables: readonly SheetTableAddition[],
): Promise<Buffer> {
  const plan = await planCellEditsToXlsx(
    await createBufferEntrySource(source),
    edits,
    structural,
    [],
    sheetPlan,
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
  return (await assembleWithJsZip(source, plan)).buffer
}

describe('sheetData anchors', () => {
  it('writes generated table headers into an added sheet that still has <sheetData/>', async () => {
    const withTable = await save(
      await blankXlsxBuffer('Sheet1'),
      [],
      [],
      { renames: [], additions: [{ name: 'T' }], removals: [], order: ['Sheet1', 'T'] },
      [
        {
          sheetName: 'T',
          area: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 2 },
          name: 'Tbl',
          columnNames: ['a', 'b', 'c'],
          bandedRows: true,
        },
      ],
    )
    const grown = await save(
      withTable,
      [],
      [{ sheetName: 'T', ops: [{ kind: 'insert-cols', index: 1, count: 1 }] }],
      undefined,
      [],
    )
    const zip = await JSZip.loadAsync(grown)
    const xml = await zip.file('xl/worksheets/sheet2.xml')!.async('string')
    expect(xml).toMatch(/<sheetData><row r="1">.*<c r="B1" t="inlineStr">/)
    expect(xml).not.toContain('<sheetData/>')
  })

  it('names the part and the bytes read when a worksheet has no sheetData at all', async () => {
    const zip = await JSZip.loadAsync(await blankXlsxBuffer('Sheet1'))
    zip.file('xl/worksheets/sheet1.xml', '')
    const broken = await zip.generateAsync({ type: 'nodebuffer' })
    await expect(
      save(
        broken,
        [{ sheetName: 'Sheet1', row: 0, column: 0, writeValue: true, cell: { value: 1 } }],
        [],
        undefined,
        [],
      ),
    ).rejects.toThrow(
      'xl/worksheets/sheet1.xml (sheet "Sheet1") has no sheetData element — 0 characters read.',
    )
  })
})
