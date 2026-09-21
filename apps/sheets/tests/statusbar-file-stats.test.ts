import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { IStatusBarService } from '@univerjs/sheets-ui'
import { SheetsSelectionsService } from '@univerjs/sheets'

import { installStatusBarFileStats, visibleRowBands } from '../src/renderer/statusbar-file-stats'

const bounds = { startRow: 0, endRow: 9, startColumn: 0, endColumn: 0 }

describe('visibleRowBands', () => {
  it('keeps a fully visible selection as one band', () => {
    expect(visibleRowBands(bounds, () => true)).toEqual([bounds])
  })

  it('splits around hidden and filtered-out rows like the Univer status bar', () => {
    const hidden = new Set([0, 3, 4, 9])
    expect(visibleRowBands(bounds, (row) => !hidden.has(row))).toEqual([
      { ...bounds, startRow: 1, endRow: 2 },
      { ...bounds, startRow: 5, endRow: 8 },
    ])
  })

  it('returns no bands when every row is hidden and null past the band cap', () => {
    expect(visibleRowBands(bounds, () => false)).toEqual([])
    expect(visibleRowBands(bounds, (row) => row % 2 === 0, 3)).toBeNull()
  })
})

type Publish = (
  state: { values: Array<{ func: string; value: number }>; pattern: string | null } | null,
) => void

function harness(opts: { hiddenRows?: Set<number>; selections?: unknown[] } = {}) {
  let publish: Publish = () => {}
  const setState = vi.fn()
  const statusBar = {
    state$: {
      subscribe(next: Publish) {
        publish = next
        return { unsubscribe: vi.fn() }
      },
    },
    setState,
  }
  const selections = opts.selections ?? [
    { range: { startRow: 0, endRow: 9, startColumn: 0, endColumn: 0, rangeType: 0 } },
  ]
  const selectionsService = { getCurrentSelections: () => selections }
  const sheet = { getRowVisible: (row: number) => !(opts.hiddenRows ?? new Set()).has(row) }
  const worksheet = {
    getSheetId: () => 's1',
    getMaxRows: () => 10,
    getMaxColumns: () => 1,
    getSheet: () => sheet,
  }
  const runtime = {
    univerAPI: { getActiveWorkbook: () => ({ getActiveSheet: () => worksheet }) },
    univer: {
      __getInjector: () => ({
        get: (token: unknown) => {
          if (token === IStatusBarService) return statusBar
          if (token === SheetsSelectionsService) return selectionsService
          throw new Error('unexpected token')
        },
      }),
    },
  }
  const state = {
    formulaMode: false,
    file: { sheets: [{ id: 's1', rowCount: 10, columnCount: 1 }] },
    editJournal: { structuralOps: new Map() },
  }
  const aggregate = vi.fn(async (_sheetId: string, b: typeof bounds) => ({
    ok: true as const,
    aggregate: {
      nonEmpty: b.endRow - b.startRow + 1,
      numericCount: b.endRow - b.startRow + 1,
      sum: b.endRow - b.startRow + 1,
      min: 1,
      max: 1,
    },
  }))
  const installed = installStatusBarFileStats({
    runtime: runtime as never,
    lazyWorkbookRef: { current: state as never },
    aggregate: aggregate as never,
  })
  return { publish: (s: Parameters<Publish>[0]) => publish(s), setState, aggregate, installed }
}

const upstream = { values: [{ func: 'COUNTA', value: 3 }], pattern: null }

describe('installStatusBarFileStats', () => {
  beforeEach(() => vi.useFakeTimers())
  afterEach(() => vi.useRealTimers())

  it('aggregates only the visible row bands of a filtered selection', async () => {
    const h = harness({ hiddenRows: new Set([2, 3, 7]) })
    h.publish(upstream)
    await vi.advanceTimersByTimeAsync(200)
    expect(h.aggregate.mock.calls.map((c) => [c[1].startRow, c[1].endRow])).toEqual([
      [0, 1],
      [4, 6],
      [8, 9],
    ])
    const values = h.setState.mock.calls[0]![0].values as Array<{ func: string; value: number }>
    expect(values.find((v) => v.func === 'COUNTA')?.value).toBe(7)
  })

  it('leaves the upstream numbers when every selected row is hidden', async () => {
    const h = harness({ hiddenRows: new Set([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]) })
    h.publish(upstream)
    await vi.advanceTimersByTimeAsync(200)
    expect(h.aggregate).not.toHaveBeenCalled()
    expect(h.setState).not.toHaveBeenCalled()
  })

  it('cancels a pending recompute when Univer clears the status bar', async () => {
    const h = harness()
    h.publish(upstream)
    await vi.advanceTimersByTimeAsync(50)
    h.publish(null)
    await vi.advanceTimersByTimeAsync(300)
    expect(h.aggregate).not.toHaveBeenCalled()
    expect(h.setState).not.toHaveBeenCalled()
    h.publish(upstream)
    await vi.advanceTimersByTimeAsync(50)
    h.publish({ values: [], pattern: null })
    await vi.advanceTimersByTimeAsync(300)
    expect(h.setState).not.toHaveBeenCalled()
  })

  it('drops an in-flight result superseded by a cleared status bar', async () => {
    const h = harness()
    let release: () => void = () => {}
    h.aggregate.mockImplementationOnce(
      () =>
        new Promise((resolve) => {
          release = () =>
            resolve({
              ok: true as const,
              aggregate: { nonEmpty: 10, numericCount: 10, sum: 10, min: 1, max: 1 },
            })
        }),
    )
    h.publish(upstream)
    await vi.advanceTimersByTimeAsync(200)
    expect(h.aggregate).toHaveBeenCalledTimes(1)
    h.publish(null)
    release()
    await vi.advanceTimersByTimeAsync(10)
    expect(h.setState).not.toHaveBeenCalled()
  })
})
