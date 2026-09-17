import { beforeEach, describe, expect, it, vi } from 'vitest'

import { installCopyMaterialize } from '../src/renderer/copy-materialize'
import { ensureLazyRangeLoaded, readCopySourceDirect } from '../src/renderer/univer-sync'

vi.mock('../src/renderer/univer-sync', () => ({
  ensureLazyRangeLoaded: vi.fn().mockResolvedValue(true),
  readCopySourceDirect: vi.fn(),
}))

const mockEnsure = vi.mocked(ensureLazyRangeLoaded)
const mockDirect = vi.mocked(readCopySourceDirect)

function harness(opts: {
  fileRows: number
  fileCols: number
  ops: unknown[]
  selection: { row: number; column: number; height: number; width: number }
  loaded?: { startRow: number; endRow: number; startColumn: number; endColumn: number }
}) {
  const messages: string[] = []
  const clipboard = { copy: vi.fn().mockResolvedValue(true), cut: vi.fn() }
  const worksheet = { getSheetId: () => 's1' }
  const selection = {
    getRow: () => opts.selection.row,
    getColumn: () => opts.selection.column,
    getHeight: () => opts.selection.height,
    getWidth: () => opts.selection.width,
  }
  const workbook = {
    getActiveSheet: () => worksheet,
    getActiveRange: () => selection,
  }
  const runtime = {
    univerAPI: { getActiveWorkbook: () => workbook },
    univer: { __getInjector: () => ({ get: () => clipboard }) },
  }
  const state = {
    formulaMode: false,
    file: { sheets: [{ id: 's1', rowCount: opts.fileRows, columnCount: opts.fileCols }] },
    editJournal: { structuralOps: new Map([['s1', opts.ops]]) },
    loadedRanges: new Map(opts.loaded ? [['s1', opts.loaded]] : []),
  }
  return {
    messages,
    clipboard,
    installer: () =>
      installCopyMaterialize(runtime as never, { current: state } as never, (message: string) =>
        messages.push(message),
      ),
  }
}

describe('copy materialize screen extent', () => {
  beforeEach(() => {
    mockEnsure.mockClear()
    mockDirect.mockReset()
  })

  it('clamps to the screen extent after row inserts, not the file extent', async () => {
    // file 10 rows + 5 inserted at top = 15 screen rows; selecting screen
    // rows 10-14 must load through row 14 (file clamping would invert to 10..9)
    const h = harness({
      fileRows: 10,
      fileCols: 8,
      ops: [{ kind: 'insert-rows', index: 0, count: 5 }],
      selection: { row: 10, column: 0, height: 5, width: 2 },
    })
    const { dispose } = h.installer()
    await h.clipboard.copy()
    expect(mockEnsure).toHaveBeenCalledTimes(1)
    const range = mockEnsure.mock.calls[0]![3] as {
      startRow: number
      endRow: number
      startColumn: number
      endColumn: number
    }
    expect(range).toMatchObject({ startRow: 10, endRow: 14, startColumn: 0, endColumn: 1 })
    dispose()
  })

  it('does nothing for selections wholly past the extent after deletes', async () => {
    // file 10 rows - 5 deleted = 5 screen rows; a stale anchor at row 20
    // inverts the range and must not reach the loader
    const h = harness({
      fileRows: 10,
      fileCols: 8,
      ops: [{ kind: 'remove-rows', index: 0, count: 5 }],
      selection: { row: 20, column: 0, height: 5, width: 2 },
    })
    const { dispose } = h.installer()
    await h.clipboard.copy()
    expect(mockEnsure).not.toHaveBeenCalled()
    expect(h.messages).toHaveLength(0)
    dispose()
  })

  it('keeps file-extent clamping when no structural ops exist', async () => {
    const h = harness({
      fileRows: 10,
      fileCols: 8,
      ops: [],
      selection: { row: 8, column: 0, height: 5, width: 2 },
    })
    const { dispose } = h.installer()
    await h.clipboard.copy()
    expect(mockEnsure).toHaveBeenCalledTimes(1)
    const range = mockEnsure.mock.calls[0]![3] as { startRow: number; endRow: number }
    expect(range).toMatchObject({ startRow: 8, endRow: 9 })
    dispose()
  })

  it('loads a selection past the old 20k cap as one resident window', async () => {
    // 1200 x 120 = 144k cells: below the full-load budget, so the grid copy
    // runs on real cells instead of the blanks the lazy window never held
    const h = harness({
      fileRows: 1200,
      fileCols: 120,
      ops: [],
      selection: { row: 0, column: 0, height: 1200, width: 120 },
    })
    const { dispose } = h.installer()
    await h.clipboard.copy()
    expect(mockEnsure).toHaveBeenCalledTimes(1)
    expect(mockEnsure.mock.calls[0]![3]).toMatchObject({
      startRow: 0,
      endRow: 1199,
      endColumn: 119,
    })
    expect(h.messages.some((m) => /A1:DP1200\b/.test(m))).toBe(true)
    dispose()
  })

  it('copies a selection past the full-load budget as values straight from the file', async () => {
    const written: string[] = []
    Object.defineProperty(globalThis, 'navigator', {
      configurable: true,
      value: { clipboard: { writeText: async (text: string) => void written.push(text) } },
    })
    mockDirect.mockResolvedValue([
      [
        { v: 'Año', t: 1, s: null, f: null, fileFormula: null },
        { v: 1234.5, t: null, s: { n: { pattern: '#,##0.00' } }, f: null, fileFormula: null },
      ],
      [
        { v: null, t: null, s: null, f: null, fileFormula: null },
        { v: 1, t: 3, s: null, f: null, fileFormula: null },
      ],
    ])
    const h = harness({
      fileRows: 12000,
      fileCols: 120,
      ops: [],
      selection: { row: 0, column: 0, height: 12000, width: 120 },
    })
    const { dispose } = h.installer()
    const ok = await h.clipboard.copy()
    expect(ok).toBe(true)
    expect(mockEnsure).not.toHaveBeenCalled()
    expect(mockDirect).toHaveBeenCalledTimes(1)
    expect(mockDirect.mock.calls[0]![2]).toMatchObject({
      startRow: 0,
      endRow: 11999,
      endColumn: 119,
    })
    expect(written).toEqual(['Año\t1,234.50\n\tTRUE'])
    expect(h.messages.at(-1)).toMatch(/A1:DP12000.*1,440,000.*250,000/)
    dispose()
  })

  it('refuses a cut past the full-load budget instead of clearing unseen cells', async () => {
    const h = harness({
      fileRows: 12000,
      fileCols: 120,
      ops: [],
      selection: { row: 0, column: 0, height: 12000, width: 120 },
    })
    const cutSpy = h.clipboard.cut
    const { dispose } = h.installer()
    expect(await h.clipboard.cut()).toBe(false)
    expect(mockDirect).not.toHaveBeenCalled()
    expect(cutSpy).not.toHaveBeenCalled()
    expect(h.messages.at(-1)).toMatch(/A1:DP12000.*250000/)
    dispose()
  })
})
