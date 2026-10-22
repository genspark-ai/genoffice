import { CellValueType } from '@univerjs/core'
import { INumfmtService, SheetsSelectionsService } from '@univerjs/sheets'
import { describe, expect, it } from 'vitest'

import { rawValueDecimals, stepSelectionDecimals } from '../src/renderer/decimal-step'
import type { UniverRuntime } from '../src/renderer/univer-state'

interface Cell {
  v?: unknown
  t?: CellValueType
  pattern?: string
}

function harness(cells: Record<string, Cell>, ranges: number[][]) {
  const cellAt = (row: number, col: number) => cells[`${row},${col}`]
  const numfmtService = {
    getValue: (_u: string, _s: string, row: number, col: number) => {
      const pattern = cellAt(row, col)?.pattern
      return pattern === undefined ? null : { pattern }
    },
  }
  const selectionsService = {
    getCurrentSelections: () =>
      ranges.map(([startRow, endRow, startColumn, endColumn]) => ({
        range: { startRow, endRow, startColumn, endColumn },
      })),
  }
  const worksheet = {
    getCellRaw: (row: number, col: number) => {
      const cell = cellAt(row, col)
      return cell?.v === undefined ? null : { v: cell.v, t: cell.t }
    },
  }
  const executed: {
    id: string
    params: { values: { row: number; col: number; pattern: string }[] }
  }[] = []
  const runtime = {
    univer: {
      __getInjector: () => ({
        get: (token: unknown) => {
          if (token === INumfmtService) return numfmtService
          if (token === SheetsSelectionsService) return selectionsService
          throw new Error('unexpected token')
        },
      }),
    },
    univerAPI: {
      getActiveWorkbook: () => ({
        getId: () => 'u1',
        getActiveSheet: () => ({
          getSheetId: () => 's1',
          getSheet: () => worksheet,
        }),
      }),
      executeCommand: (id: string, params: never) => {
        executed.push({ id, params })
        return Promise.resolve(true)
      },
    },
  } as unknown as UniverRuntime
  return { runtime, executed }
}

const values = (executed: ReturnType<typeof harness>['executed']) => executed[0]?.params.values

describe('rawValueDecimals', () => {
  it('counts decimal digits of number cells only', () => {
    expect(rawValueDecimals({ v: 341504.77, t: CellValueType.NUMBER })).toBe(2)
    expect(rawValueDecimals({ v: 7751938, t: CellValueType.NUMBER })).toBe(0)
    expect(rawValueDecimals({ v: '341504.77', t: CellValueType.STRING })).toBe(0)
    expect(rawValueDecimals(null)).toBe(0)
  })
})

describe('stepSelectionDecimals', () => {
  it('treats an explicit General record like no record: one click adds ONE place', () => {
    // The regression: Univer reads getFormatInfo('General').maxDecimals = 9,
    // so one Increase Decimal click produced a 10-place pattern.
    const h = harness({ '0,0': { v: 341504.77, t: CellValueType.NUMBER, pattern: 'General' } }, [
      [0, 0, 0, 0],
    ])
    stepSelectionDecimals(h.runtime, 1)
    expect(values(h.executed)).toEqual([{ row: 0, col: 0, pattern: '0.000' }])
  })

  it('gives a decimal-less General number a single place', () => {
    const h = harness({ '0,0': { v: 7751938, t: CellValueType.NUMBER } }, [[0, 0, 0, 0]])
    stepSelectionDecimals(h.runtime, 1)
    expect(values(h.executed)).toEqual([{ row: 0, col: 0, pattern: '0.0' }])
  })

  it('steps an explicit numeric pattern by one place in both directions', () => {
    const up = harness({ '0,0': { v: 1.5, t: CellValueType.NUMBER, pattern: '0.00' } }, [
      [0, 0, 0, 0],
    ])
    stepSelectionDecimals(up.runtime, 1)
    expect(values(up.executed)).toEqual([{ row: 0, col: 0, pattern: '0.000' }])
    const down = harness({ '0,0': { v: 1.5, t: CellValueType.NUMBER, pattern: '0.00' } }, [
      [0, 0, 0, 0],
    ])
    stepSelectionDecimals(down.runtime, -1)
    expect(values(down.executed)).toEqual([{ row: 0, col: 0, pattern: '0.0' }])
  })

  it('seeds General cells from the largest raw decimal count in the selection', () => {
    const h = harness(
      {
        '0,0': { v: 341504.77, t: CellValueType.NUMBER, pattern: 'General' },
        '1,0': { v: 2000000, t: CellValueType.NUMBER },
      },
      [[0, 1, 0, 0]],
    )
    stepSelectionDecimals(h.runtime, 1)
    expect(values(h.executed)).toEqual([
      { row: 0, col: 0, pattern: '0.000' },
      { row: 1, col: 0, pattern: '0.000' },
    ])
  })

  it('decrease on a decimal-less General cell yields a plain integer pattern', () => {
    const h = harness({ '0,0': { v: 7751938, t: CellValueType.NUMBER } }, [[0, 0, 0, 0]])
    stepSelectionDecimals(h.runtime, -1)
    expect(values(h.executed)).toEqual([{ row: 0, col: 0, pattern: '0' }])
  })

  it('increase skips cells whose pattern cannot take decimals', () => {
    const h = harness({ '0,0': { v: 'label', t: CellValueType.STRING, pattern: '@' } }, [
      [0, 0, 0, 0],
    ])
    stepSelectionDecimals(h.runtime, 1)
    expect(h.executed).toHaveLength(0)
  })
})
