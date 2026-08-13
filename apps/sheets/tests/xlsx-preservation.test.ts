import { describe, expect, it } from 'vitest'

import { applyPlanToXlsx, inventoryXlsx, readBasicWorkbook } from '../src/gateway/xlsx-gateway'
import type { ChangePlan } from '../src/domain/workbook.types'
import { buildCompatibilityFixture } from './fixture-builder'

describe('XLSX preservation gateway', () => {
  it('imports worksheet values for the desktop editor', async () => {
    const imported = await readBasicWorkbook(await buildCompatibilityFixture())

    expect(imported.snapshot.sheets).toEqual([
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: {
          A1: { value: 'Old' },
          B1: { value: 10 },
        },
      },
    ])
    expect(imported.sheetNamesById).toEqual({ 'sheet-1': 'Sheet1' })
  })

  it('changes only the target worksheet entry', async () => {
    const source = await buildCompatibilityFixture()
    const mutation = await applyPlanToXlsx(source, createPlan(), { 'sheet-1': 'Sheet1' })
    const before = new Map(mutation.beforeEntries.map((entry) => [entry.path, entry.sha256]))
    const after = new Map(mutation.afterEntries.map((entry) => [entry.path, entry.sha256]))

    expect(mutation.touchedEntries).toEqual(['xl/worksheets/sheet1.xml'])
    expect(after.get('xl/worksheets/sheet1.xml')).not.toBe(before.get('xl/worksheets/sheet1.xml'))
    for (const [path, hash] of before) {
      if (path !== 'xl/worksheets/sheet1.xml') expect(after.get(path)).toBe(hash)
    }
  })

  it('writes a formula while preserving unknown and chart parts', async () => {
    const source = await buildCompatibilityFixture()
    const formulaPlan: ChangePlan = {
      ...createPlan(),
      transactionId: 'formula-1',
      cellChanges: [
        {
          sheetId: 'sheet-1',
          address: 'B2',
          before: { value: null },
          after: { value: null, formula: '=SUM(B1:B1)' },
        },
      ],
    }
    const mutation = await applyPlanToXlsx(source, formulaPlan, { 'sheet-1': 'Sheet1' })
    const entries = await inventoryXlsx(mutation.buffer)
    const paths = entries.map((entry) => entry.path)

    expect(paths).toContain('customXml/item1.xml')
    expect(paths).toContain('xl/charts/chart1.xml')
  })

  it('fails closed when a sheet mapping is absent', async () => {
    const source = await buildCompatibilityFixture()
    await expect(applyPlanToXlsx(source, createPlan(), {})).rejects.toThrow(
      'Missing XLSX sheet mapping',
    )
  })

  it('rejects a write when the package content differs from the preview', async () => {
    const source = await buildCompatibilityFixture()
    const stalePlan: ChangePlan = {
      ...createPlan(),
      cellChanges: [
        {
          sheetId: 'sheet-1',
          address: 'A1',
          before: { value: 'Unexpected' },
          after: { value: 'New' },
        },
      ],
    }
    await expect(applyPlanToXlsx(source, stalePlan, { 'sheet-1': 'Sheet1' })).rejects.toThrow(
      'no longer has the expected content',
    )
  })
})

function createPlan(): ChangePlan {
  return {
    transactionId: 'tx-1',
    baseRevision: 0,
    cellChanges: [
      {
        sheetId: 'sheet-1',
        address: 'A1',
        before: { value: 'Old' },
        after: { value: 'New' },
      },
    ],
    sheetRenames: [],
    structuralChanges: [],
    formatChanges: [],
    warnings: [],
  }
}
