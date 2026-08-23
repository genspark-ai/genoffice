import { describe, expect, it } from 'vitest'

import { handleRibbonCommand, type RibbonCommandContext } from '../src/renderer/ribbon-actions'

function makeHarness() {
  const calls: string[] = []
  const selection = { startRow: 0, endRow: 9, startColumn: 0, endColumn: 4 }
  const worksheet = { getSheetId: () => 's1' }
  const workbook = {
    getId: () => 'wb',
    getActiveSheet: () => worksheet,
    getActiveRange: () => ({ getRange: () => selection }),
  }
  const skeleton = {
    rowHeightAccumulation: Array.from({ length: 100 }, (_, index) => (index + 1) * 20),
    columnWidthAccumulation: Array.from({ length: 20 }, (_, index) => (index + 1) * 80),
    rowHeaderWidthAndMarginLeft: 46,
    columnHeaderHeightAndMarginTop: 24,
  }
  const render = {
    engine: { width: 846, height: 424 },
    with: () => ({ getCurrentSkeleton: () => skeleton }),
  }
  const univerAPI = {
    getActiveWorkbook: () => workbook,
    executeCommand: async (id: string) => {
      calls.push(id)
      return true
    },
  }
  const ctx = {
    univerRef: {
      current: {
        univerAPI,
        univer: { __getInjector: () => ({ get: () => ({ getRenderById: () => render }) }) },
      },
    },
    setMessage: () => {},
  } as unknown as RibbonCommandContext
  return { ctx, calls }
}

describe('zoom-to-selection ordering', () => {
  it('scrolls only after the zoom command resolves', async () => {
    const { ctx, calls } = makeHarness()
    handleRibbonCommand(ctx, 'zoom-to-selection')
    // Synchronously only the zoom is dispatched; scrolling at the old ratio
    // clamps short near the grid edge.
    expect(calls).toEqual(['sheet.command.set-zoom-ratio'])
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(calls).toEqual(['sheet.command.set-zoom-ratio', 'sheet.command.scroll-to-cell'])
  })
})
