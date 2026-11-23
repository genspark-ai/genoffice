import { describe, expect, it } from 'vitest'

import { snapshotLiveWorkbook } from '../src/renderer/collab-snapshot'
import type { UniverRuntime } from '../src/renderer/univer-state'

function firstSheet(snapshot: ReturnType<typeof snapshotLiveWorkbook>) {
  const sheet = snapshot?.sheets[0]
  if (!sheet) throw new Error('the snapshot has no sheet')
  return sheet
}

/**
 * The room publishes the live grid, not the AI file model. These tests pin the
 * fields a peer needs to see a manual edit: values, formulas, resolved styles,
 * merges and row/column sizes.
 */
function fakeRuntime(): UniverRuntime {
  const worksheet = {
    getRange: (address: string) => ({
      getNumberFormat: () => (address === 'A1' ? 'General' : '0.00'),
    }),
    getSheetName: () => 'Sheet1',
    getMaxRows: () => 1000,
    getMaxColumns: () => 26,
  }
  const workbook = {
    getSnapshot: () => ({
      sheetOrder: ['sheet-1'],
      styles: { s1: { fs: 14, bl: 1, cl: { rgb: '#FF0000' } }, s2: { fs: 11 } },
      sheets: {
        'sheet-1': {
          id: 'sheet-1',
          name: 'Sheet1',
          rowCount: 1000,
          columnCount: 26,
          cellData: {
            0: { 0: { v: 'hello', s: 's1' }, 1: { f: '=A1&"!"', s: 's2' } },
          },
          mergeData: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }],
          rowData: { 0: { h: 48 } },
          columnData: { 1: { w: 120 } },
        },
      },
    }),
    getSheetBySheetId: (id: string) => (id === 'sheet-1' ? worksheet : null),
  }
  return {
    univerAPI: { getActiveWorkbook: () => workbook },
  } as unknown as UniverRuntime
}

describe('snapshotLiveWorkbook', () => {
  it('reads the live grid, not the AI file model', () => {
    const snapshot = snapshotLiveWorkbook(fakeRuntime(), 7)
    expect(snapshot).not.toBeNull()
    expect(snapshot!.revision).toBe(7)
    expect(snapshot!.sheets).toHaveLength(1)
    const sheet = firstSheet(snapshot)
    expect(sheet.id).toBe('sheet-1')
    expect(sheet.name).toBe('Sheet1')
    expect(sheet.cells.A1).toEqual({ value: 'hello' })
    expect(sheet.cells.B1).toEqual({ value: null, formula: '=A1&"!"' })
  })

  it('resolves styles, merges and row/column sizes', () => {
    const sheet = firstSheet(snapshotLiveWorkbook(fakeRuntime(), 7))
    expect(sheet.styles!.A1).toMatchObject({
      bold: true,
      fontSize: 14,
      fontColor: '#FF0000',
    })
    expect(sheet.styles!.A1).not.toHaveProperty('numberFormat')
    expect(sheet.styles!.B1).toMatchObject({ numberFormat: '0.00' })
    expect(sheet.merges).toEqual(['A1:B1'])
    expect(sheet.rowHeights).toEqual({ 1: 36 })
    expect(sheet.colWidths).toEqual({ B: 120 })
  })

  it('carries only chart visuals for the matching sheet', () => {
    const visuals = [
      { id: 'chart-1', sheetId: 'sheet-1', kind: 'chart' },
      { id: 'pic-1', sheetId: 'sheet-1', kind: 'image' },
      { id: 'chart-2', sheetId: 'sheet-2', kind: 'chart' },
    ] as never
    const sheet = firstSheet(snapshotLiveWorkbook(fakeRuntime(), 7, visuals))
    expect(sheet.visuals).toEqual([{ id: 'chart-1', sheetId: 'sheet-1', kind: 'chart' }])
  })

  it('returns null without a live workbook', () => {
    expect(snapshotLiveWorkbook(null, 1)).toBeNull()
  })
})
