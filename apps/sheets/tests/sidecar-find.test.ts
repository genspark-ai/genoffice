import { afterEach, describe, expect, it, vi } from 'vitest'

import type { WorkbookFindCellsResult } from '../src/shared/desktop-api'
import {
  findSheetCellsInFile,
  sidecarFindAvailable,
  type SidecarFindHit,
} from '../src/renderer/sidecar-find'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

const QUERY = {
  query: 'x',
  matchCase: false,
  matchEntireCell: false,
  lookIn: 'formulas',
  wildcards: false,
} as const

function state(structuralOps: unknown[] = [], aliases?: Map<string, string>): LazyWorkbookState {
  return {
    file: {
      sessionId: '11111111-1111-4111-8111-111111111111',
      sheets: [{ id: 'file-1', name: 'Data', rowCount: 1000, columnCount: 10 }],
    },
    streamAliases: aliases,
    editJournal: { cells: new Map(), structuralOps: new Map([['file-1', structuralOps]]) },
  } as unknown as LazyWorkbookState
}

function page(
  matches: { row: number; column: number; value?: unknown; formula?: string }[],
  rest: Partial<WorkbookFindCellsResult> = {},
): WorkbookFindCellsResult {
  return {
    matches: matches.map((m) => ({
      sheetId: 'file-1',
      row: m.row,
      column: m.column,
      value: (m.value ?? 'x') as string,
      valueText: m.value === undefined ? 'x' : String(m.value),
      ...(m.formula === undefined ? {} : { formulaText: m.formula }),
    })),
    complete: true,
    indexingComplete: true,
    ...rest,
  }
}

function install(findWorkbookCells: ReturnType<typeof vi.fn>) {
  vi.stubGlobal('window', { desktopApi: { findWorkbookCells } })
  return findWorkbookCells
}

afterEach(() => {
  vi.unstubAllGlobals()
})

describe('sidecarFindAvailable', () => {
  it('is false without the preload API', () => {
    expect(sidecarFindAvailable()).toBe(false)
    vi.stubGlobal('window', { desktopApi: {} })
    expect(sidecarFindAvailable()).toBe(false)
    install(vi.fn())
    expect(sidecarFindAvailable()).toBe(true)
  })
})

describe('findSheetCellsInFile', () => {
  it('follows nextCursor across pages and emits each page', async () => {
    const api = install(
      vi
        .fn()
        .mockResolvedValueOnce(
          page([{ row: 1, column: 0 }], {
            complete: false,
            nextCursor: { sheetId: 'file-1', row: 1, column: 1 },
          }),
        )
        .mockResolvedValueOnce(page([{ row: 7, column: 2, formula: '=X1' }])),
    )
    const pages: SidecarFindHit[][] = []
    const outcome = await findSheetCellsInFile(state(), 'file-1', QUERY, {
      onPage: (hits) => pages.push(hits) && hits.length,
    })
    expect(outcome).toEqual({ complete: true, capped: false, matches: 2 })
    expect(pages).toEqual([
      [{ row: 1, column: 0, value: 'x', valueText: 'x', formula: undefined }],
      [{ row: 7, column: 2, value: 'x', valueText: 'x', formula: '=X1' }],
    ])
    expect(api.mock.calls[0]![0]).toMatchObject({
      sessionId: '11111111-1111-4111-8111-111111111111',
      sheetId: 'file-1',
      ...QUERY,
      limit: 100_000,
    })
    expect(api.mock.calls[1]![0]).toMatchObject({
      resumeAt: { sheetId: 'file-1', row: 1, column: 1 },
    })
  })

  it('maps file coordinates through journaled structural ops and drops deleted lines', async () => {
    install(
      vi.fn().mockResolvedValue(
        page([
          { row: 0, column: 0 },
          { row: 2, column: 0 },
          { row: 5, column: 3 },
        ]),
      ),
    )
    const ops = [
      { kind: 'insert-rows', index: 1, count: 2 },
      { kind: 'remove-rows', index: 4, count: 1 },
      { kind: 'insert-cols', index: 0, count: 1 },
    ]
    const hits: SidecarFindHit[] = []
    await findSheetCellsInFile(state(ops), 'file-1', QUERY, {
      onPage: (batch) => hits.push(...batch) && batch.length,
    })
    // File row 2 → screen 4 (two inserted rows above), then deleted.
    expect(hits.map((hit) => [hit.row, hit.column])).toEqual([
      [0, 1],
      [6, 4],
    ])
  })

  it('resolves stream aliases to the file sheet id', async () => {
    const api = install(vi.fn().mockResolvedValue(page([])))
    await findSheetCellsInFile(state([], new Map([['copy-1', 'file-1']])), 'copy-1', QUERY, {
      onPage: (hits) => hits.length,
    })
    expect(api.mock.calls[0]![0]).toMatchObject({ sheetId: 'file-1' })
  })

  it('stops at maxMatches and reports the cap', async () => {
    const api = install(
      vi.fn().mockResolvedValue(
        page(
          [
            { row: 0, column: 0 },
            { row: 0, column: 1 },
          ],
          { complete: false, nextCursor: { sheetId: 'file-1', row: 0, column: 2 } },
        ),
      ),
    )
    const outcome = await findSheetCellsInFile(state(), 'file-1', QUERY, {
      maxMatches: 2,
      onPage: (hits) => hits.length,
    })
    expect(outcome).toEqual({ complete: false, capped: true, matches: 2 })
    expect(api).toHaveBeenCalledTimes(1)
    expect(api.mock.calls[0]![0]).toMatchObject({ limit: 2 })
  })

  it('retries while indexing lags and honors the wait limit', async () => {
    const lagging = page([], {
      complete: false,
      indexingComplete: false,
      nextCursor: { sheetId: 'file-1', row: 256, column: 0 },
    })
    const api = install(
      vi
        .fn()
        .mockResolvedValueOnce(lagging)
        .mockResolvedValueOnce(page([{ row: 300, column: 0 }])),
    )
    const hits: SidecarFindHit[] = []
    const outcome = await findSheetCellsInFile(state(), 'file-1', QUERY, {
      onPage: (batch) => hits.push(...batch) && batch.length,
    })
    expect(outcome.complete).toBe(true)
    expect(hits).toHaveLength(1)
    expect(api).toHaveBeenCalledTimes(2)

    install(vi.fn().mockResolvedValue(lagging))
    const gaveUp = await findSheetCellsInFile(state(), 'file-1', QUERY, {
      indexingWaitLimitMs: 0,
      onPage: (hits) => hits.length,
    })
    expect(gaveUp).toEqual({ complete: false, capped: false, matches: 0 })
  })

  it('rejects on a failed first page but keeps earlier pages on a later failure', async () => {
    install(vi.fn().mockRejectedValue(new Error('unsupported command')))
    await expect(
      findSheetCellsInFile(state(), 'file-1', QUERY, { onPage: (hits) => hits.length }),
    ).rejects.toThrow('unsupported command')

    install(
      vi
        .fn()
        .mockResolvedValueOnce(
          page([{ row: 0, column: 0 }], {
            complete: false,
            nextCursor: { sheetId: 'file-1', row: 0, column: 1 },
          }),
        )
        .mockRejectedValueOnce(new Error('timed out')),
    )
    const outcome = await findSheetCellsInFile(state(), 'file-1', QUERY, {
      onPage: (hits) => hits.length,
    })
    expect(outcome).toEqual({ complete: false, capped: false, matches: 1 })
  })

  it('stops between pages once the caller is gone', async () => {
    const api = install(
      vi.fn().mockResolvedValue(
        page([{ row: 0, column: 0 }], {
          complete: false,
          nextCursor: { sheetId: 'file-1', row: 0, column: 1 },
        }),
      ),
    )
    let alive = true
    const outcome = await findSheetCellsInFile(state(), 'file-1', QUERY, {
      alive: () => alive,
      onPage: (hits) => {
        alive = false
        return hits.length
      },
    })
    expect(outcome.complete).toBe(false)
    expect(api).toHaveBeenCalledTimes(1)
  })

  it('skips sheets without file backing', async () => {
    const api = install(vi.fn())
    const outcome = await findSheetCellsInFile(state(), 'added-this-session', QUERY, {
      onPage: (hits) => hits.length,
    })
    expect(outcome).toEqual({ complete: true, capped: false, matches: 0 })
    expect(api).not.toHaveBeenCalled()
  })
})

describe('findSheetCellsInFile cap accounting', () => {
  afterEach(() => {
    vi.unstubAllGlobals()
  })

  it('counts only the hits the caller kept toward maxMatches', async () => {
    const api = install(
      vi
        .fn()
        .mockResolvedValueOnce(
          page(
            [
              { row: 0, column: 0 },
              { row: 0, column: 1 },
            ],
            { complete: false, nextCursor: { sheetId: 'file-1', row: 0, column: 2 } },
          ),
        )
        .mockResolvedValueOnce(page([{ row: 9, column: 0 }])),
    )
    const outcome = await findSheetCellsInFile(state(), 'file-1', QUERY, {
      maxMatches: 2,
      // the first page's hits are all journal-shadowed: nothing kept
      onPage: (hits) => (hits[0]!.row === 0 ? 0 : hits.length),
    })
    expect(outcome).toEqual({ complete: true, capped: false, matches: 1 })
    expect(api).toHaveBeenCalledTimes(2)
  })
})
