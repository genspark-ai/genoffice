import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { IRange } from '@univerjs/core'
import { FindModel, type IFindMatch } from '@univerjs/find-replace'
import { Subject } from 'rxjs'

import {
  buildLazyCellTest,
  collectJournalMatches,
  coveredByWindow,
  extraComparator,
  installLazyFindBridge,
  journalShadowKeys,
  mergeFindMatches,
  planLazyFind,
  scalarToText,
  type LazyCellMatch,
} from '../src/renderer/lazy-find'
import { ensureLazyRangeLoaded, readSheetRangeMapped } from '../src/renderer/univer-sync'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

vi.mock('../src/renderer/univer-sync', () => ({
  readSheetRangeMapped: vi.fn(),
  ensureLazyRangeLoaded: vi.fn().mockResolvedValue(true),
}))

vi.mock('../src/renderer/i18n/locale', () => ({
  t: (key: string, params?: Record<string, string>) =>
    `${key}${params ? ` ${JSON.stringify(params)}` : ''}`,
}))

const mockRead = vi.mocked(readSheetRangeMapped)
const mockEnsure = vi.mocked(ensureLazyRangeLoaded)

describe('planLazyFind', () => {
  it('is inactive without a streamed workbook', () => {
    expect(planLazyFind(null)).toBe('inactive')
  })

  it('is inactive once preloading completed', () => {
    expect(planLazyFind(state({ preloadComplete: true }))).toBe('inactive')
  })

  it('extends the search while rows are still streaming', () => {
    expect(planLazyFind(state({ preloadComplete: false }))).toBe('extend')
  })
})

describe('scalarToText', () => {
  it('stringifies like Univer does', () => {
    expect(scalarToText(120)).toBe('120')
    expect(scalarToText(true)).toBe('1')
    expect(scalarToText(false)).toBe('0')
    expect(scalarToText('txt')).toBe('txt')
    expect(scalarToText(null)).toBeNull()
    expect(scalarToText(undefined)).toBeNull()
  })
})

describe('buildLazyCellTest', () => {
  it('matches substrings case-insensitively by default', () => {
    const test = buildLazyCellTest(query({ findString: 'Total' }))
    expect(test?.({ value: 'grand TOTAL', formula: undefined })).toBe(true)
    expect(test?.({ value: 'other', formula: undefined })).toBe(false)
  })

  it('honors case sensitivity', () => {
    const test = buildLazyCellTest(query({ findString: 'Total', caseSensitive: true }))
    expect(test?.({ value: 'grand total', formula: undefined })).toBe(false)
    expect(test?.({ value: 'grand Total', formula: undefined })).toBe(true)
  })

  it('trims spaces (not line breaks) for whole-cell matches', () => {
    const test = buildLazyCellTest(query({ findString: 'total', matchesTheWholeCell: true }))
    expect(test?.({ value: '  total  ', formula: undefined })).toBe(true)
    // Line breaks are kept, mirroring Univer's trimLeadingTrailingWhitespace.
    expect(test?.({ value: 'total\n', formula: undefined })).toBe(false)
    expect(test?.({ value: 'grand total', formula: undefined })).toBe(false)
  })

  it('looks at formulas only when searching formulas', () => {
    const formulaQuery = query({ findString: 'a2*2', findBy: 'formula' })
    expect(formulaQuery.findBy).toBe('formula')
    const test = buildLazyCellTest(formulaQuery)
    expect(test?.({ value: '240', formula: '=A2*2' })).toBe(true)
    expect(test?.({ value: '240', formula: undefined })).toBe(false)
    const valueTest = buildLazyCellTest(query({ findString: '240', findBy: 'value' }))
    expect(valueTest?.({ value: '240', formula: '=A2*2' })).toBe(true)
  })

  it('trims the needle like the built-in preprocessor', () => {
    const test = buildLazyCellTest(query({ findString: '  Total  ' }))!
    expect(test({ value: 'subtotal', formula: undefined })).toBe(true)
    expect(buildLazyCellTest(query({ findString: '   ' }))).toBeNull()
  })

  it('rejects an empty needle', () => {
    expect(buildLazyCellTest(query({ findString: '' }))).toBeNull()
  })
})

describe('mergeFindMatches', () => {
  const inner = match('s1', 5, 1)
  const extraSame = match('s1', 5, 1)
  const extraFar = match('s1', 900, 0)

  it('keeps inner entries and appends unseen extras', () => {
    const merged = mergeFindMatches([inner], [extraSame, extraFar])
    expect(merged).toHaveLength(2)
    expect(merged[0]).toBe(inner)
    expect(merged[1]).toBe(extraFar)
  })

  it('returns primary untouched without extras', () => {
    const merged = mergeFindMatches([inner], [])
    expect(merged).toEqual([inner])
  })
})

describe('coveredByWindow', () => {
  it('checks the sheet loaded range', () => {
    const lazyState = state({})
    expect(coveredByWindow(lazyState, 's1', 5, 5)).toBe(true)
    expect(coveredByWindow(lazyState, 's1', 50, 0)).toBe(false)
    expect(coveredByWindow(lazyState, 'ghost', 0, 0)).toBe(false)
  })
})

describe('collectJournalMatches', () => {
  it('skips loaded-window edits and cleared cells', () => {
    const lazyState = state({
      journalCells: new Map([
        [
          's1',
          new Map([
            ['0', { row: 5, column: 0, value: 'inside hit', hasValue: true }],
            ['1', { row: 900, column: 0, value: 'outside hit', hasValue: true }],
            ['2', { row: 901, column: 0, value: null, hasValue: false }],
          ]),
        ],
      ]),
    })
    const found = collectJournalMatches(
      lazyState,
      's1',
      buildLazyCellTest(query({ findString: 'hit' }))!,
    )
    expect(found.map((cell) => cell.value)).toEqual(['outside hit'])
  })

  it('tests formulas of journaled cells', () => {
    const lazyState = state({
      journalCells: new Map([
        [
          's1',
          new Map([
            ['0', { row: 20, column: 2, value: null, formula: '=Sum(A1:A9)', hasValue: true }],
          ]),
        ],
      ]),
    })
    const found = collectJournalMatches(
      lazyState,
      's1',
      buildLazyCellTest(query({ findString: 'sum(', findBy: 'formula' }))!,
    )
    expect(found).toHaveLength(1)
    expect(found[0]?.formula).toBe('=Sum(A1:A9)')
  })
})

describe('extraComparator', () => {
  const order = new Map([
    ['s1', 0],
    ['s2', 1],
  ])
  const cell = (sheetId: string, row: number, column: number) => ({
    sheetId,
    row,
    column,
    value: null,
    formula: undefined,
  })

  it('orders row-major across sheets', () => {
    const compare = extraComparator(order, false)
    expect(compare(cell('s1', 1, 9), cell('s2', 0, 0))).toBeLessThan(0)
    expect(compare(cell('s1', 2, 0), cell('s1', 1, 5))).toBeGreaterThan(0)
    expect(compare(cell('s1', 1, 2), cell('s1', 1, 2))).toBe(0)
  })

  it('orders column-major when requested', () => {
    const compare = extraComparator(order, true)
    expect(compare(cell('s1', 9, 1), cell('s1', 0, 2))).toBeLessThan(0)
  })
})

type FakeOverrides = {
  preloadComplete?: boolean
  journalCells?: LazyWorkbookState['editJournal']['cells']
  rowCount?: number
}

function state(overrides: FakeOverrides): LazyWorkbookState {
  return {
    file: {
      sessionId: 'session-1',
      sheets: [{ id: 's1', name: 'Sheet1', rowCount: overrides.rowCount ?? 1000, columnCount: 8 }],
    },
    generation: 1,
    loadedRanges: new Map<string, IRange>([
      ['s1', { startRow: 0, endRow: 9, startColumn: 0, endColumn: 9 }],
    ]),
    loadingKeys: new Map(),
    retryTimers: new Map(),
    appliedMerges: new Map(),
    appliedRowKeys: new Map(),
    sheetProtections: new Map(),
    sheetPageBreaks: new Map(),
    sheetProtectedRanges: new Map(),
    uninstalledDefinedNames: new Set(),
    appliedCfSheets: new Set(),
    appliedFilterSheets: new Set(),
    appliedDvSheets: new Set(),
    decorationsPendingSheets: new Set(),
    hyperlinkTargets: new Map(),
    frozenStripKeys: new Map(),
    filterOrigins: new Map(),
    showFormulaSheets: new Set(),
    formulaMode: false,
    editJournal: {
      cells: overrides.journalCells ?? new Map(),
      structuralOps: new Map(),
    },
    flags: { preloadComplete: overrides.preloadComplete ?? false },
    closure: { status: 'idle', pinned: new Map() },
    formulaText: new Map(),
    cachedFormulaValues: new Map(),
    pivotDefinitions: new Map(),
    outline: new Map(),
    recalc: {
      timer: null,
      generation: 0,
      failures: 0,
      formulaCells: new Map(),
      overlay: new Map(),
    },
  } as unknown as LazyWorkbookState
}

function query(overrides: Record<string, unknown> = {}): ReturnType<typeof buildQuery> {
  return buildQuery(overrides)
}

function buildQuery(overrides: Record<string, unknown>) {
  return {
    findString: 'needle',
    caseSensitive: false,
    findBy: 'value',
    findDirection: 'row',
    findScope: 'subunit',
    matchesTheWholeCell: false,
    replaceRevealed: false,
    ...overrides,
  } as Parameters<typeof buildLazyCellTest>[0]
}

function match(sheetId: string, row: number, column: number): LazyCellMatch {
  return {
    provider: 'sheets-find-replace-provider',
    unitId: 'workbook-1',
    isFormula: false,
    range: {
      subUnitId: sheetId,
      range: { startRow: row, endRow: row, startColumn: column, endColumn: column },
    },
  }
}

class FakeInnerModel extends FindModel {
  readonly unitId = 'workbook-1'
  readonly matchesUpdate$ = new Subject<IFindMatch[]>()
  readonly activelyChangingMatch$ = new Subject<LazyCellMatch>()

  constructor(private readonly matches: IFindMatch[]) {
    super()
  }

  getMatches(): IFindMatch[] {
    return this.matches
  }

  moveToNextMatch(): IFindMatch | null {
    return this.matches[0] ?? null
  }

  moveToPreviousMatch(): IFindMatch | null {
    return this.matches[this.matches.length - 1] ?? null
  }

  replace(_replaceString: string): Promise<boolean> {
    return Promise.resolve(false)
  }

  async replaceAll(): Promise<{ success: number; failure: number }> {
    return { success: this.matches.length, failure: 0 }
  }

  focusSelection(): void {}
}

/** Index-cursor navigation like the built-in SheetFindModel: currentMatch
 *  advances by list index, loop takes the modulo, exhaustion resets the
 *  cursor, ignoreSelection starts from the first/last match. */
class CursorInnerModel extends FakeInnerModel {
  private cursor = -1
  /** Like the built-in, focusSelection moves the grid selection to the
   *  current match. */
  onFocus: ((row: number, column: number) => void) | null = null

  constructor(private readonly list: IFindMatch[]) {
    super(list)
  }

  override focusSelection(): void {
    if (this.cursor < 0 || !this.onFocus) return
    const bounds = (this.list[this.cursor] as LazyCellMatch).range.range
    this.onFocus(bounds.startRow, bounds.startColumn)
  }

  override moveToNextMatch(params?: { loop?: boolean }): IFindMatch | null {
    return this.step(1, params)
  }

  override moveToPreviousMatch(params?: { loop?: boolean }): IFindMatch | null {
    return this.step(-1, params)
  }

  private step(delta: number, params?: { loop?: boolean }): IFindMatch | null {
    if (this.list.length === 0) return null
    if (this.cursor < 0) {
      this.cursor = delta > 0 ? 0 : this.list.length - 1
      return this.list[this.cursor]!
    }
    const next = this.cursor + delta
    if (next < 0 || next >= this.list.length) {
      if (params?.loop) {
        this.cursor = (next + this.list.length) % this.list.length
        return this.list[this.cursor]!
      }
      this.cursor = -1
      return null
    }
    this.cursor = next
    return this.list[next]!
  }
}

function facade(lazyState: LazyWorkbookState | null) {
  const setValues = vi.fn()
  // Selection follows activation, like the live grid does.
  const active = { row: 0, column: 0 }
  const worksheet = {
    getSheetId: () => 's1',
    getSheetName: () => 'Sheet1',
    scrollToCell: vi.fn(),
    getRange: vi.fn((row: number, column: number) => ({
      activate: () => {
        active.row = row
        active.column = column
      },
      setValues,
    })),
  }
  const workbook = {
    getId: () => 'workbook-1',
    getSheets: () => [worksheet],
    getActiveSheet: () => worksheet,
    getSheetBySheetId: (sheetId: string) => (sheetId === 's1' ? worksheet : null),
    getActiveRange: () => ({ getRow: () => active.row, getColumn: () => active.column }),
    setActiveSheet: vi.fn(),
  }
  const providers = new Set<unknown>()
  const registrations: { provider: unknown; dispose: () => void }[] = []
  const service = {
    getProviders: () => providers,
    registerFindReplaceProvider: (provider: unknown) => {
      const registration = {
        provider,
        dispose: () => {
          providers.delete(registration.provider)
        },
      }
      providers.add(provider)
      registrations.push(registration)
      return registration
    },
  }
  const runtime = {
    univerAPI: { getActiveWorkbook: () => workbook },
    univer: { __getInjector: () => ({ get: () => service }) },
  }
  return {
    runtime: runtime as unknown as Parameters<typeof installLazyFindBridge>[0]['runtime'],
    lazyWorkbookRef: { current: lazyState },
    setMessage: vi.fn(),
    worksheet,
    workbook,
    setValues,
    providers,
    service,
    registrations,
    active,
  }
}

async function settle(model: { getMatches(): unknown[] }): Promise<void> {
  await vi.waitFor(() => {
    expect(mockRead).toHaveBeenCalled()
    expect(model.getMatches().length).toBeGreaterThan(0)
  })
}

type MappedResult = Awaited<ReturnType<typeof readSheetRangeMapped>>

function mapped(
  cells: {
    row: number
    column: number
    value: string | number | boolean | null
    formula?: string
  }[],
): MappedResult {
  return {
    raw: { indexingComplete: true },
    indexedThroughScreen: 999,
    fileEndRow: 999,
    screen: { cells, rows: [], merges: [], hyperlinks: [] },
  } as unknown as MappedResult
}

describe('installLazyFindBridge', () => {
  beforeEach(() => {
    mockRead.mockReset()
    mockEnsure.mockReset()
    mockEnsure.mockResolvedValue(true)
  })

  it('passes through when no streamed workbook is open', async () => {
    const harness = facade(null)
    const inner = new FakeInnerModel([match('s1', 1, 1)])
    const builtin = { find: vi.fn().mockResolvedValue([inner]), terminate: vi.fn() }
    harness.providers.add(builtin)
    const bridge = installLazyFindBridge(harness)

    const models = await harnessLookup(harness)(query())
    expect(models).toHaveLength(1)
    expect(models[0]).toBe(inner)
    bridge.dispose()
    // The wrapper left, the built-in stayed registered.
    expect([...harness.providers]).toContain(builtin)
  })

  it('extends the session with out-of-window hits and focuses them', async () => {
    const harness = facade(state({}))
    const inner = new FakeInnerModel([])
    const builtin = { find: vi.fn().mockResolvedValue([inner]), terminate: vi.fn() }
    harness.providers.add(builtin)
    const bridge = installLazyFindBridge(harness)

    mockRead.mockResolvedValue(
      mapped([
        { row: 500, column: 3, value: 'deep needle' },
        { row: 5, column: 3, value: 'window needle' },
      ]),
    )

    const models = await harnessLookup(harness)(query())
    expect(models).toHaveLength(1)
    const model = models[0]!
    await settle(model)

    const matches = model.getMatches()
    // The in-window file hit belongs to the (empty) inner list; only the deep one is added.
    expect(matches).toHaveLength(1)
    expect((matches[0] as LazyCellMatch).range.range.startRow).toBe(500)

    const focused = model.moveToNextMatch()
    expect(focused).not.toBeNull()
    expect(harness.worksheet.scrollToCell).toHaveBeenCalledWith(500, 3)
    expect(mockEnsure).toHaveBeenCalled()
    bridge.dispose()
  })

  it('dedupes hits the inner model also reports', async () => {
    const harness = facade(state({}))
    const inner = new FakeInnerModel([match('s1', 500, 3)])
    const builtin = { find: vi.fn().mockResolvedValue([inner]), terminate: vi.fn() }
    harness.providers.add(builtin)
    const bridge = installLazyFindBridge(harness)

    mockRead.mockResolvedValue(mapped([{ row: 500, column: 3, value: 'same needle' }]))

    const models = await harnessLookup(harness)(query())
    const model = models[0]!
    await settle(model)
    expect(model.getMatches()).toHaveLength(1)
    bridge.dispose()
  })

  it('reports truncated scans through the status message', async () => {
    const harness = facade(state({}))
    const inner = new FakeInnerModel([])
    const builtin = { find: vi.fn().mockResolvedValue([inner]), terminate: vi.fn() }
    harness.providers.add(builtin)
    const bridge = installLazyFindBridge(harness)

    mockRead.mockRejectedValue(new Error('sidecar gone'))

    const models = await harnessLookup(harness)(query())
    const model = models[0]!
    await vi.waitFor(() => expect(harness.setMessage).toHaveBeenCalled())
    expect(model.getMatches()).toHaveLength(0)
    // A failed first read scanned nothing — the message must say so instead
    // of quoting the budget cap.
    expect(harness.setMessage.mock.calls[0]![0]).toContain('"cells":"0"')
    bridge.dispose()
  })

  it('stops scanning by scanned extent, not by hit count', async () => {
    // 1M rows x 8 columns = 8M cells; the 400k budget must cut the scan off
    // after ~23 of the ~445 batches even though nothing matches.
    const harness = facade(state({ rowCount: 1_000_000 }))
    const inner = new FakeInnerModel([])
    const builtin = { find: vi.fn().mockResolvedValue([inner]), terminate: vi.fn() }
    harness.providers.add(builtin)
    const bridge = installLazyFindBridge(harness)

    mockRead.mockResolvedValue(mapped([]))

    await harnessLookup(harness)(query())
    await vi.waitFor(() => expect(harness.setMessage).toHaveBeenCalled())
    expect(mockRead.mock.calls.length).toBeLessThan(30)
    const message = harness.setMessage.mock.calls[0]![0] as string
    expect(message).toContain('appFindScanTruncated')
    // Actual scanned extent (414k), not the 400,000 constant.
    expect(message).toMatch(/414/)
    bridge.dispose()
  })

  it('walks past in-window hits into the extras and wraps globally', async () => {
    const harness = facade(state({}))
    const inner = new CursorInnerModel([match('s1', 1, 1), match('s1', 2, 2)])
    inner.onFocus = (row, column) => {
      harness.active.row = row
      harness.active.column = column
    }
    const builtin = { find: vi.fn().mockResolvedValue([inner]), terminate: vi.fn() }
    harness.providers.add(builtin)
    const bridge = installLazyFindBridge(harness)

    mockRead.mockResolvedValue(mapped([{ row: 500, column: 3, value: 'deep needle' }]))

    const models = await harnessLookup(harness)(query())
    const model = models[0]!
    await vi.waitFor(() => expect(model.getMatches()).toHaveLength(3))

    // The dialog's Find Next always passes loop: true on a single model.
    const move = () => model.moveToNextMatch({ loop: true }) as LazyCellMatch | null
    expect(move()!.range.range.startRow).toBe(1)
    expect(move()!.range.range.startRow).toBe(2)
    expect(move()!.range.range.startRow).toBe(500)
    expect(harness.worksheet.scrollToCell).toHaveBeenCalledWith(500, 3)
    // Extras exhausted: wrap back into the inner session, full cycle.
    expect(move()!.range.range.startRow).toBe(1)
    expect(move()!.range.range.startRow).toBe(2)
    expect(move()!.range.range.startRow).toBe(500)
    bridge.dispose()
  })

  it('keeps file matches findable under style-only journal edits', async () => {
    const journalCells = new Map([
      [
        's1',
        new Map([
          ['500:3', { row: 500, column: 3, value: null, hasValue: false, style: { bold: true } }],
          ['600:3', { row: 600, column: 3, value: 'overwritten', hasValue: true }],
        ]),
      ],
    ]) as unknown as LazyWorkbookState['editJournal']['cells']
    const lazyState = state({ journalCells })
    expect(journalShadowKeys(lazyState, 's1')).toEqual(new Set(['600:3']))

    const harness = facade(lazyState)
    const inner = new FakeInnerModel([])
    const builtin = { find: vi.fn().mockResolvedValue([inner]), terminate: vi.fn() }
    harness.providers.add(builtin)
    const bridge = installLazyFindBridge(harness)

    mockRead.mockResolvedValue(
      mapped([
        { row: 500, column: 3, value: 'file needle under style edit' },
        { row: 600, column: 3, value: 'file needle overwritten in session' },
      ]),
    )

    const models = await harnessLookup(harness)(query())
    const model = models[0]!
    await settle(model)
    const rows = model.getMatches().map((hit) => (hit as LazyCellMatch).range.range.startRow)
    // The style-only edit leaves the file value findable; the overwrite hides it.
    expect(rows).toEqual([500])
    bridge.dispose()
  })

  it('follows column-major order for by-column searches', async () => {
    const harness = facade(state({}))
    const inner = new FakeInnerModel([])
    const builtin = { find: vi.fn().mockResolvedValue([inner]), terminate: vi.fn() }
    harness.providers.add(builtin)
    const bridge = installLazyFindBridge(harness)

    mockRead.mockResolvedValue(
      mapped([
        { row: 100, column: 2, value: 'needle a' },
        { row: 500, column: 1, value: 'needle b' },
      ]),
    )

    const models = await harnessLookup(harness)(query({ findDirection: 'column' }))
    const model = models[0]!
    await vi.waitFor(() => expect(model.getMatches()).toHaveLength(2))

    const first = model.moveToNextMatch() as LazyCellMatch
    expect(first.range.range.startColumn).toBe(1)
    expect(first.range.range.startRow).toBe(500)
    bridge.dispose()
  })

  it('drops the extra cursor once the inner session resumes', async () => {
    const harness = facade(state({}))
    const inner = new CursorInnerModel([match('s1', 1, 1)])
    inner.onFocus = (row, column) => {
      harness.active.row = row
      harness.active.column = column
    }
    const builtin = { find: vi.fn().mockResolvedValue([inner]), terminate: vi.fn() }
    harness.providers.add(builtin)
    const bridge = installLazyFindBridge(harness)

    mockRead.mockResolvedValue(mapped([{ row: 500, column: 3, value: 'deep needle' }]))

    const models = await harnessLookup(harness)(query())
    const model = models[0]!
    await vi.waitFor(() => expect(model.getMatches()).toHaveLength(2))

    const move = () => model.moveToNextMatch({ loop: true }) as LazyCellMatch | null
    expect(move()!.range.range.startRow).toBe(1)
    expect(move()!.range.range.startRow).toBe(500)
    expect(move()!.range.range.startRow).toBe(1)

    // Inner owns the cursor again: neither focus nor replace may touch the
    // stale extra.
    harness.worksheet.scrollToCell.mockClear()
    model.focusSelection()
    expect(harness.worksheet.scrollToCell).not.toHaveBeenCalled()
    await model.replace('x')
    expect(harness.setValues).not.toHaveBeenCalled()
    bridge.dispose()
  })

  it('does not touch the grid once the workbook state went stale', async () => {
    const harness = facade(state({}))
    const inner = new FakeInnerModel([])
    const builtin = { find: vi.fn().mockResolvedValue([inner]), terminate: vi.fn() }
    harness.providers.add(builtin)
    const bridge = installLazyFindBridge(harness)

    mockRead.mockResolvedValue(mapped([{ row: 500, column: 3, value: 'deep needle' }]))

    const models = await harnessLookup(harness)(query())
    const model = models[0]!
    await settle(model)

    harness.lazyWorkbookRef.current = null
    model.moveToNextMatch()
    expect(harness.worksheet.scrollToCell).not.toHaveBeenCalled()
    expect(mockEnsure).not.toHaveBeenCalled()
    bridge.dispose()
  })
})

/** A model whose disposal is observable, like the built-in SheetFindModel. */
class DisposableInnerModel extends FakeInnerModel {
  disposed = false

  override dispose(): void {
    this.disposed = true
    super.dispose()
  }
}

/** Mimics SheetsFindReplaceProvider: find() first terminates prior models. */
function univerLikeBuiltin(matches: IFindMatch[]) {
  let liveModels: DisposableInnerModel[] = []
  const terminate = () => {
    liveModels.forEach((model) => model.dispose())
    liveModels = []
  }
  return {
    findCalls: 0,
    async find(_query: unknown) {
      this.findCalls += 1
      terminate()
      const model = new DisposableInnerModel(matches)
      liveModels = [model]
      return [model]
    },
    terminate,
    liveModel: () => liveModels[0] ?? null,
  }
}

/** Mirrors FindReplaceModel._startSearching: dispatches to EVERY provider in
 *  the service's live set — the semantics that made append-only registration
 *  double-run the built-in. */
async function dispatchLikeUniver(harness: ReturnType<typeof facade>): Promise<FindModel[]> {
  const list = Array.from(harness.providers) as {
    find(q: unknown): Promise<FindModel[]>
  }[]
  return (await Promise.all(list.map((provider) => provider.find(query())))).flat()
}

describe('service-level dispatch (Univer semantics)', () => {
  beforeEach(() => {
    mockRead.mockReset()
    mockEnsure.mockReset()
    mockEnsure.mockResolvedValue(true)
  })

  it('detaches the built-in so a search reaches it exactly once', async () => {
    const harness = facade(state({}))
    const builtin = univerLikeBuiltin([match('s1', 1, 1)])
    harness.providers.add(builtin)
    const bridge = installLazyFindBridge(harness)
    expect([...harness.providers]).toHaveLength(1)

    mockRead.mockResolvedValue(mapped([]))
    const models = await dispatchLikeUniver(harness)
    expect(builtin.findCalls).toBe(1)
    expect(models).toHaveLength(1)
    expect(models[0]!.getMatches()).toHaveLength(1)
    expect(builtin.liveModel()!.disposed).toBe(false)

    bridge.dispose()
    expect([...harness.providers]).toContain(builtin)
  })

  it('sweeps up a built-in that registers after the bridge', async () => {
    const harness = facade(state({}))
    const bridge = installLazyFindBridge(harness)
    const builtin = univerLikeBuiltin([match('s1', 1, 1)])
    harness.service.registerFindReplaceProvider(builtin)

    mockRead.mockResolvedValue(mapped([]))
    await dispatchLikeUniver(harness)
    expect([...harness.providers]).not.toContain(builtin)

    const before = builtin.findCalls
    const models = await dispatchLikeUniver(harness)
    expect(builtin.findCalls).toBe(before + 1)
    expect(models).toHaveLength(1)
    expect(models[0]!.getMatches()).toHaveLength(1)
    bridge.dispose()
  })
})

/** Runs a find through the provider the bridge registered (last registration). */
function harnessLookup(harness: ReturnType<typeof facade>): (q: unknown) => Promise<FindModel[]> {
  const registration = harness.registrations[harness.registrations.length - 1]!
  const wrapper = registration.provider as { find: (q: unknown) => Promise<FindModel[]> }
  expect(typeof wrapper.find).toBe('function')
  return (q) => wrapper.find(q)
}
