import { describe, expect, it } from 'vitest'
import type { MutablePackage } from '../src/gateway/xlsx-drawing-add'
import {
  applyTableAdditions,
  TableAddError,
  type TableAddition,
} from '../src/gateway/xlsx-table-add'
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from '../src/shared/grid-bounds'

const emptyPackage: MutablePackage = {
  paths: async () => [],
  has: async () => false,
  readText: async () => {
    throw new Error('validation must reject before any part is read')
  },
  write: () => {},
  add: () => {},
  addBinary: () => {},
  remove: () => {},
}

function addition(overrides: Partial<TableAddition>): TableAddition {
  return {
    worksheetPath: 'xl/worksheets/sheet1.xml',
    area: { startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 },
    name: 'Table1',
    columnNames: ['A', 'B'],
    bandedRows: true,
    ...overrides,
  }
}

async function reject(overrides: Partial<TableAddition>): Promise<void> {
  await expect(applyTableAdditions(emptyPackage, [addition(overrides)], new Set())).rejects.toThrow(
    TableAddError,
  )
}

describe('table addition validation', () => {
  it('rejects names that break the defined-name rules', async () => {
    await reject({ name: 'My Table' })
    await reject({ name: '1Table' })
    await reject({ name: 'A1' })
    await reject({ name: 'R1C1' })
    await reject({ name: 'x'.repeat(256) })
    await reject({ name: '' })
  })

  it('rejects areas outside the worksheet grid', async () => {
    await reject({ area: { startRow: 0, startColumn: 0, endRow: MAX_GRID_ROWS, endColumn: 1 } })
    await reject({
      area: {
        startRow: 0,
        startColumn: MAX_GRID_COLUMNS - 1,
        endRow: 2,
        endColumn: MAX_GRID_COLUMNS,
      },
    })
    await reject({ area: { startRow: -1, startColumn: 0, endRow: 2, endColumn: 1 } })
    await reject({ area: { startRow: 3, startColumn: 0, endRow: 2, endColumn: 1 } })
  })
})
