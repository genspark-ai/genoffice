import { describe, expect, it } from 'vitest'

import { loadWorkbookSkeleton, normalizeVisibleRange } from '../src/renderer/univer-sync'

/// loadWorkbookSkeleton clears the unit's undo history through the injector.
const undoStub = {
  __getInjector: () => ({ get: () => ({ clearUndoRedo: () => undefined }) }),
}

describe('normalizeVisibleRange', () => {
  it('uses the initial viewport when Univer has no scroll range yet', () => {
    expect(normalizeVisibleRange(null, 14_516, 16)).toEqual({
      startRow: 0,
      endRow: 79,
      startColumn: 0,
      endColumn: 15,
    })
  })

  it('falls back when a stale viewport is outside the replacement sheet', () => {
    expect(
      normalizeVisibleRange(
        {
          startRow: 20_000,
          endRow: 20_050,
          startColumn: 20,
          endColumn: 25,
        },
        14_516,
        16,
      ),
    ).toEqual({
      startRow: 0,
      endRow: 79,
      startColumn: 0,
      endColumn: 15,
    })
  })

  it('clamps a partially out-of-bounds viewport', () => {
    expect(
      normalizeVisibleRange(
        {
          startRow: -5,
          endRow: 200,
          startColumn: -2,
          endColumn: 30,
        },
        100,
        16,
      ),
    ).toEqual({
      startRow: 0,
      endRow: 99,
      startColumn: 0,
      endColumn: 15,
    })
  })
})

describe('loadWorkbookSkeleton', () => {
  it('shows a full starter grid for a blank A1-only workbook', () => {
    const created: Array<{
      sheets: Record<string, { rowCount: number; columnCount: number }>
    }> = []
    const runtime = {
      univer: undoStub,
      univerAPI: {
        getActiveWorkbook: () => null,
        disposeUnit: () => undefined,
        createWorkbook: (config: (typeof created)[number]) => {
          created.push(config)
          return { getSheetBySheetId: () => null, setActiveSheet: () => undefined }
        },
      },
    }
    const file = {
      sha256: 'blank',
      name: 'Book1.xlsx',
      visuals: [],
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet1',
          rowCount: 1,
          columnCount: 1,
          hidden: false,
          showGridLines: true,
          tabColor: null,
          defaultRowHeight: null,
          defaultColumnWidth: null,
          freeze: null,
          columnWidths: [],
        },
      ],
    }

    loadWorkbookSkeleton(runtime as never, file as never)

    expect(created[0]?.sheets['sheet-1']).toMatchObject({
      rowCount: 1000,
      columnCount: 26,
    })
  })

  it('preserves a file extent larger than the starter grid', () => {
    const created: Array<{
      sheets: Record<string, { rowCount: number; columnCount: number }>
    }> = []
    const runtime = {
      univer: undoStub,
      univerAPI: {
        getActiveWorkbook: () => null,
        disposeUnit: () => undefined,
        createWorkbook: (config: (typeof created)[number]) => {
          created.push(config)
          return { getSheetBySheetId: () => null, setActiveSheet: () => undefined }
        },
      },
    }
    const file = {
      sha256: 'large',
      name: 'Large.xlsx',
      visuals: [],
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet1',
          rowCount: 1004,
          columnCount: 13,
          hidden: false,
          showGridLines: true,
          tabColor: null,
          defaultRowHeight: null,
          defaultColumnWidth: null,
          freeze: null,
          columnWidths: [],
        },
      ],
    }

    loadWorkbookSkeleton(runtime as never, file as never)

    expect(created[0]?.sheets['sheet-1']).toMatchObject({
      rowCount: 1004,
      columnCount: 26,
    })
  })
})
