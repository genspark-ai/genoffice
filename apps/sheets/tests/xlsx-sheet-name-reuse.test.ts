import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { applyCellEditsToXlsx, type CellEdit } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { SheetEditError } from '@genoffice/xlsx-gateway/gateway/xlsx-sheets'
import { buildSheetsFixture } from './fixture-builder'

// Per-sheet states name an existing sheet by its original file name and an
// added sheet by its final name. Renaming "Data" and adding a new "Data" in one
// save made both read as one sheet: a value typed on the renamed sheet landed
// on the new one. Such a save is refused before any state is applied when a
// state names that sheet.
const RENAME_AND_REUSE = {
  renames: [{ sheetName: 'Data', newName: 'Numbers' }],
  additions: [{ name: 'Data' }],
  removals: [],
  order: ['Numbers', 'Other', 'My Sheet', 'Data'],
}
const VALUE: CellEdit = {
  sheetName: 'Data',
  row: 9,
  column: 0,
  writeValue: true,
  cell: { value: 'meant-for-renamed' },
}

async function sheetText(buffer: Buffer, sheetName: string): Promise<string> {
  const zip = await JSZip.loadAsync(buffer)
  const workbook = (await zip.file('xl/workbook.xml')?.async('text')) ?? ''
  const rels = (await zip.file('xl/_rels/workbook.xml.rels')?.async('text')) ?? ''
  const id = new RegExp(`<sheet\\b[^>]*name="${sheetName}"[^>]*r:id="([^"]+)"`).exec(workbook)?.[1]
  const target = new RegExp(`Id="${id}"[^>]*Target="([^"]+)"`).exec(rels)?.[1]
  return (await zip.file(`xl/${target}`)?.async('text')) ?? ''
}

async function sheetNames(buffer: Buffer): Promise<string[]> {
  const zip = await JSZip.loadAsync(buffer)
  const workbook = (await zip.file('xl/workbook.xml')?.async('text')) ?? ''
  return [...workbook.matchAll(/<sheet\b[^>]*name="([^"]+)"/g)].map((match) => match[1]!)
}

describe('a new sheet that takes the old name of a sheet renamed in the same save', () => {
  const cases: [string, () => Promise<unknown>][] = [
    [
      'a cell value',
      async () =>
        applyCellEditsToXlsx(await buildSheetsFixture(), [VALUE], [], [], RENAME_AND_REUSE),
    ],
    [
      'a row operation',
      async () =>
        applyCellEditsToXlsx(
          await buildSheetsFixture(),
          [],
          [{ sheetName: 'Data', ops: [{ kind: 'insert-rows', index: 0, count: 1 }] }],
          [],
          RENAME_AND_REUSE,
        ),
    ],
    [
      'conditional formatting',
      async () =>
        applyCellEditsToXlsx(
          await buildSheetsFixture(),
          [],
          [],
          [],
          RENAME_AND_REUSE,
          [],
          [],
          [{ sheetName: 'Data', rules: [] }],
        ),
    ],
    [
      'data validation',
      async () =>
        applyCellEditsToXlsx(
          await buildSheetsFixture(),
          [],
          [],
          [],
          RENAME_AND_REUSE,
          [],
          [],
          [],
          [{ sheetName: 'Data', rules: [] }],
        ),
    ],
    [
      'page setup and print areas',
      async () =>
        applyCellEditsToXlsx(
          await buildSheetsFixture(),
          [],
          [],
          [],
          RENAME_AND_REUSE,
          [],
          [],
          [],
          [],
          [],
          null,
          [
            { sheetName: 'Data', printArea: 'A1:C10', orientation: 'landscape' },
            { sheetName: 'Data', printArea: 'A1:B2' },
          ],
        ),
    ],
    [
      'a visibility change in the plan',
      async () =>
        applyCellEditsToXlsx(await buildSheetsFixture(), [], [], [], {
          ...RENAME_AND_REUSE,
          hiddenChanges: [{ sheetName: 'Data', hidden: true }],
        }),
    ],
  ]
  for (const [label, save] of cases) {
    it(`refuses the save with ${label}`, async () => {
      const error = await save().catch((caught: unknown) => caught)
      expect(error).toBeInstanceOf(SheetEditError)
      expect((error as Error).message).toMatch(/a new sheet takes its old name/)
    })
  }

  it('also refuses when the old name was freed by a chained rename', async () => {
    await expect(
      applyCellEditsToXlsx(await buildSheetsFixture(), [{ ...VALUE, sheetName: 'Other' }], [], [], {
        renames: [
          { sheetName: 'Data', newName: 'X' },
          { sheetName: 'Other', newName: 'Data' },
        ],
        additions: [{ name: 'Other' }],
        removals: [],
        order: ['X', 'Data', 'My Sheet', 'Other'],
      }),
    ).rejects.toThrow(/a new sheet takes its old name/)
  })
})

describe('sheet plans that stay unambiguous', () => {
  // Nothing names either sheet, so nothing can land on the wrong one; the
  // CLI's `sheet apply` saves this batch as well (packages/cli xlsx-mixed-batch).
  it('saves the rename and the new sheet when no state names them', async () => {
    const mutation = await applyCellEditsToXlsx(
      await buildSheetsFixture(),
      [],
      [],
      [],
      RENAME_AND_REUSE,
    )
    expect(await sheetNames(mutation.buffer)).toEqual(['Numbers', 'Other', 'My Sheet', 'Data'])
  })

  it('saves a value on a renamed sheet next to a new sheet with another name', async () => {
    const mutation = await applyCellEditsToXlsx(await buildSheetsFixture(), [VALUE], [], [], {
      renames: [{ sheetName: 'Data', newName: 'Numbers' }],
      additions: [{ name: 'Fresh' }],
      removals: [],
      order: ['Numbers', 'Other', 'My Sheet', 'Fresh'],
    })
    expect(await sheetText(mutation.buffer, 'Numbers')).toContain('meant-for-renamed')
    expect(await sheetText(mutation.buffer, 'Fresh')).not.toContain('meant-for-renamed')
  })

  it('saves a swap of two names, with the value on the sheet it was typed on', async () => {
    const mutation = await applyCellEditsToXlsx(await buildSheetsFixture(), [VALUE], [], [], {
      renames: [
        { sheetName: 'Data', newName: 'Other' },
        { sheetName: 'Other', newName: 'Data' },
      ],
      additions: [],
      removals: [],
      order: ['Other', 'Data', 'My Sheet'],
    })
    expect(await sheetText(mutation.buffer, 'Other')).toContain('meant-for-renamed')
    expect(await sheetText(mutation.buffer, 'Data')).not.toContain('meant-for-renamed')
  })

  // A removed sheet carries no states (the renderer drops them), so a new sheet
  // reusing its name is the only sheet that name can mean.
  it('saves a new sheet that takes a removed sheet name, with its values on the new sheet', async () => {
    const mutation = await applyCellEditsToXlsx(
      await buildSheetsFixture(),
      [{ ...VALUE, sheetName: 'Other' }],
      [],
      [],
      {
        renames: [],
        additions: [{ name: 'Other' }],
        removals: ['Other'],
        order: ['Data', 'My Sheet', 'Other'],
      },
    )
    expect(await sheetNames(mutation.buffer)).toEqual(['Data', 'My Sheet', 'Other'])
    expect(await sheetText(mutation.buffer, 'Other')).toContain('meant-for-renamed')
  })
})
