import { ICommandService } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import { journalExternalLinkLiterals } from '../src/renderer/univer-sync'
import type { ExternalLinkUsageCell } from '../src/shared/desktop-api'

/// A formula cell as the grid holds it: the cached value the literal takes,
/// plus the shared-formula/style bits the rollback has to put back.
type FakeCell = Record<string, unknown>

function fakeRuntime(): {
  runtime: never
  matrix: Record<number, Record<number, FakeCell>>
  executed: { command: string; params: Record<string, unknown> }[]
} {
  const matrix: Record<number, Record<number, FakeCell>> = {
    0: {
      0: { f: '[1]Sheet1!A1', v: 20, t: 2 },
      1: { f: '[1]Sheet1!B1', v: 'gone!', t: 1 },
    },
  }
  const executed: { command: string; params: Record<string, unknown> }[] = []
  const sheet = {
    getSheetName: () => 'Sheet1',
    getSheetId: () => 'sheet-1',
    getSheet: () => ({ getCellMatrix: () => ({ getMatrix: () => matrix }) }),
  }
  const workbook = { getId: () => 'unit-1', getSheets: () => [sheet] }
  const runtime = {
    univerAPI: { getActiveWorkbook: () => workbook },
    univer: {
      __getInjector: () => ({
        get: (token: unknown) =>
          token === ICommandService
            ? {
                executeCommand: (command: string, params: Record<string, unknown>) => {
                  executed.push({ command, params })
                  return Promise.resolve(true)
                },
              }
            : undefined,
      }),
    },
  }
  return { runtime: runtime as never, matrix, executed }
}

const CELLS: ExternalLinkUsageCell[] = [
  { sheet: 'Sheet1', row: 0, column: 0, value: { kind: 'number', value: 20 } },
  { sheet: 'Sheet1', row: 0, column: 1, value: { kind: 'text', value: 'gone!' } },
]

describe('journalExternalLinkLiterals', () => {
  it('writes the literals and reports how many cells it journaled', async () => {
    const { runtime, executed } = fakeRuntime()
    const journal = await journalExternalLinkLiterals(runtime, CELLS)
    expect(journal.journaled).toBe(2)
    expect(executed).toHaveLength(1)
    expect(executed[0]!.command).toBe('sheet.command.set-range-values')
    const value = executed[0]!.params.value as Record<number, Record<number, FakeCell>>
    expect(value[0]![0]).toEqual({ v: 20, t: 2, f: null, si: null, p: null })
    expect(value[0]![1]).toEqual({ v: 'gone!', t: 1, f: null, si: null, p: null })
  })

  it('keeps the previous formulas for the failure rollback', async () => {
    const { runtime, executed } = fakeRuntime()
    const journal = await journalExternalLinkLiterals(runtime, CELLS)
    await journal.restore()
    expect(executed).toHaveLength(2)
    const previous = executed[1]!.params.value as Record<number, Record<number, FakeCell>>
    expect(previous[0]![0]).toEqual({ f: '[1]Sheet1!A1', v: 20, t: 2 })
    expect(previous[0]![1]).toEqual({ f: '[1]Sheet1!B1', v: 'gone!', t: 1 })
  })

  it('does not journal cells whose cache is an error or missing', async () => {
    const { runtime, executed } = fakeRuntime()
    const journal = await journalExternalLinkLiterals(runtime, [
      { sheet: 'Sheet1', row: 0, column: 0, value: { kind: 'error', value: '#REF!' } },
      { sheet: 'Sheet1', row: 0, column: 0, value: { kind: 'missing' } },
    ])
    expect(journal.journaled).toBe(0)
    expect(executed).toHaveLength(0)
  })
})
