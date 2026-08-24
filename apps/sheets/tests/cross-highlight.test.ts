import { afterEach, describe, expect, it, vi } from 'vitest'

import {
  CROSS_HIGHLIGHT_CANVAS_COLORS,
  CrossHighlightExtension,
  crossHighlightRects,
  installCrossHighlight,
  loadCrossHighlightPreference,
  resolveCrossHighlightTarget,
} from '../src/renderer/cross-highlight'

const cellAt = (row: number, column: number) => ({
  startX: column * 10,
  endX: (column + 1) * 10,
  startY: row * 20,
  endY: (row + 1) * 20,
})

describe('crossHighlightRects', () => {
  it('covers the active row and column across the visible grid', () => {
    expect(
      crossHighlightRects(4, 2, [{ startRow: 0, endRow: 9, startColumn: 0, endColumn: 5 }], cellAt),
    ).toEqual([
      { key: 'row', left: 0, top: 80, width: 60, height: 20 },
      { key: 'column', left: 20, top: 0, width: 10, height: 200 },
    ])
  })

  it('continues across frozen-pane viewport segments', () => {
    expect(
      crossHighlightRects(
        4,
        3,
        [
          { startRow: 0, endRow: 1, startColumn: 2, endColumn: 5 },
          { startRow: 2, endRow: 9, startColumn: 0, endColumn: 1 },
          { startRow: 2, endRow: 9, startColumn: 2, endColumn: 5 },
        ],
        cellAt,
      ),
    ).toEqual([
      { key: 'column', left: 30, top: 0, width: 10, height: 40 },
      { key: 'row', left: 0, top: 80, width: 20, height: 20 },
      { key: 'row', left: 20, top: 80, width: 40, height: 20 },
      { key: 'column', left: 30, top: 40, width: 10, height: 160 },
    ])
  })

  it('supports blank cells at large coordinates without an extent cap', () => {
    const rects = crossHighlightRects(
      50_000,
      3_000,
      [{ startRow: 49_990, endRow: 50_010, startColumn: 2_990, endColumn: 3_010 }],
      cellAt,
    )
    expect(rects).toEqual([
      { key: 'row', left: 29_900, top: 1_000_000, width: 210, height: 20 },
      { key: 'column', left: 30_000, top: 999_800, width: 10, height: 420 },
    ])
  })

  it('rejects invalid active coordinates', () => {
    const range = [{ startRow: 0, endRow: 9, startColumn: 0, endColumn: 5 }]
    expect(crossHighlightRects(-1, 0, range, cellAt)).toEqual([])
    expect(crossHighlightRects(0, -1, range, cellAt)).toEqual([])
    expect(crossHighlightRects(Number.NaN, 0, range, cellAt)).toEqual([])
  })
})

describe('resolveCrossHighlightTarget', () => {
  it('uses the primary active cell rather than the selection bounds', () => {
    expect(
      resolveCrossHighlightTarget({
        getId: () => 'book-1',
        getActiveSheet: () => ({ getSheetId: () => 'sheet-2' }),
        getActiveCell: () => ({ getRow: () => 8, getColumn: () => 5 }),
      }),
    ).toEqual({ workbookId: 'book-1', sheetId: 'sheet-2', row: 8, column: 5 })
  })
})

describe('loadCrossHighlightPreference', () => {
  afterEach(() => vi.unstubAllGlobals())

  it('reads the stored flag', () => {
    const store = new Map<string, string>([['ai-sheets-cross-highlight', '1']])
    vi.stubGlobal('window', { localStorage: { getItem: (key: string) => store.get(key) ?? null } })
    expect(loadCrossHighlightPreference()).toBe(true)
    store.set('ai-sheets-cross-highlight', '0')
    expect(loadCrossHighlightPreference()).toBe(false)
  })

  it('defaults to off without a localStorage', () => {
    // No window at all (node test env): must not throw.
    expect(loadCrossHighlightPreference()).toBe(false)
  })
})

describe('installCrossHighlight', () => {
  it('repaints without allocating per-selection float DOM resources', () => {
    let row = 1
    let column = 2
    let theme: 'light' | 'dark' = 'light'
    const handlers = new Map<string, () => void>()
    const disposeSelection = vi.fn()
    const disposeSheet = vi.fn()
    const makeComponentDirty = vi.fn()
    const makeSceneDirty = vi.fn()
    const registerComponent = vi.fn()
    const addFloatDomToRange = vi.fn()
    const workbook = {
      getId: () => 'book-1',
      getActiveSheet: () => ({ getSheetId: () => 'sheet-1', addFloatDomToRange }),
      getActiveCell: () => ({ getRow: () => row, getColumn: () => column }),
    }
    const runtime = {
      univerAPI: {
        Event: {
          SelectionChanged: 'selection-changed',
          ActiveSheetChanged: 'active-sheet-changed',
        },
        getActiveWorkbook: () => workbook,
        addEvent: vi.fn((event: string, handler: () => void) => {
          handlers.set(event, handler)
          return {
            dispose: event === 'selection-changed' ? disposeSelection : disposeSheet,
          }
        }),
        registerComponent,
      },
      univer: {
        __getInjector: () => ({
          get: () => ({
            getRenderById: () => ({
              mainComponent: { makeDirty: makeComponentDirty },
              scene: { makeDirty: makeSceneDirty },
            }),
          }),
        }),
      },
    } as unknown as Parameters<typeof installCrossHighlight>[0]

    const handle = installCrossHighlight(runtime, { theme: () => theme })
    handle.setVisible(true)
    for (let index = 0; index < 50; index += 1) {
      row += 1
      column += 1
      handlers.get('selection-changed')?.()
    }
    theme = 'dark'
    handle.refresh()

    expect(makeComponentDirty).toHaveBeenCalledTimes(52)
    expect(makeSceneDirty).toHaveBeenCalledTimes(52)
    expect(registerComponent).not.toHaveBeenCalled()
    expect(addFloatDomToRange).not.toHaveBeenCalled()
    expect(CROSS_HIGHLIGHT_CANVAS_COLORS.dark.fill).toContain('0.12')

    const fillRect = vi.fn()
    const canvas = {
      __mode: 'normal',
      fillStyle: '',
      strokeStyle: '',
      lineWidth: 1,
      save: vi.fn(),
      fillRect,
      getScale: () => ({ scaleX: 1, scaleY: 1 }),
      beginPath: vi.fn(),
      moveTo: vi.fn(),
      lineTo: vi.fn(),
      stroke: vi.fn(),
      closePath: vi.fn(),
      restore: vi.fn(),
    }
    const skeleton = {
      worksheet: { unitId: 'book-1', getSheetId: () => 'sheet-1' },
      rowColumnSegment: { startRow: 50, endRow: 60, startColumn: 50, endColumn: 60 },
      getCellWithCoordByIndex: cellAt,
    }
    new CrossHighlightExtension().draw(
      canvas as never,
      { scaleX: 1, scaleY: 1 },
      skeleton as never,
      undefined,
      {
        viewRanges: [{ startRow: 50, endRow: 60, startColumn: 50, endColumn: 60 }],
        viewportKey: 'VIEW_MAIN',
      },
    )
    expect(fillRect).toHaveBeenCalledTimes(2)
    expect(canvas.fillStyle).toBe(CROSS_HIGHLIGHT_CANVAS_COLORS.dark.fill)

    handle.dispose()
    expect(disposeSelection).toHaveBeenCalledTimes(1)
    expect(disposeSheet).toHaveBeenCalledTimes(1)
  })
})
