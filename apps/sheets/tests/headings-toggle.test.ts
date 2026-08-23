import { describe, expect, it } from 'vitest'

import { BooleanNumber } from '@univerjs/core'

import { handleRibbonCommand, type RibbonCommandContext } from '../src/renderer/ribbon-actions'

/// Replays the viewport math of Univer's set-row-header-width handler,
/// including its bug: the shift comes from `cornerViewport.width || 46`,
/// which reads 46 after a hide, so re-showing shifts the grid by zero.
function makeHeadingsHarness() {
  const config = {
    rowHeader: { hidden: BooleanNumber.FALSE },
    columnHeader: { hidden: BooleanNumber.FALSE },
  }
  const skeleton = { rowHeaderWidth: 46 }
  const viewLeftTop = { width: 46 }
  const viewMain = { left: 46 }
  const viewColumnRight = {
    left: 46,
    setViewportSize(props: { left?: number }) {
      if (props.left !== undefined) this.left = props.left
    },
  }
  const scene = {
    getViewport: (key: string) =>
      (({ viewMain, viewColumnRight, viewLeftTop }) as Record<string, unknown>)[key],
    makeDirty: () => {},
  }
  const render = { scene, with: () => ({ getCurrentSkeleton: () => skeleton }) }
  const worksheet = {
    getSheetId: () => 'sheet1',
    getSheet: () => ({ getConfig: () => config }),
  }
  const workbook = { getId: () => 'wb1', getActiveSheet: () => worksheet }
  const univerAPI = {
    getActiveWorkbook: () => workbook,
    executeCommand: async (id: string, params: { size: number }) => {
      if (id === 'sheet.command.set-row-header-width') {
        skeleton.rowHeaderWidth = params.size
        const delta = params.size - (viewLeftTop.width || 46)
        viewMain.left += delta
        viewColumnRight.setViewportSize({ left: viewColumnRight.left + delta })
      }
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
    lazyWorkbookRef: { current: null },
    setMessage: () => {},
    setPendingEdits: () => {},
  } as unknown as RibbonCommandContext
  return { ctx, config, viewMain, viewColumnRight }
}

const settle = () => new Promise((resolve) => setTimeout(resolve, 0))

describe('toggle-headings viewport offsets', () => {
  it('hides and re-shows without drifting the grid', async () => {
    const { ctx, config, viewMain, viewColumnRight } = makeHeadingsHarness()

    handleRibbonCommand(ctx, 'toggle-headings')
    await settle()
    expect(config.rowHeader.hidden).toBe(BooleanNumber.TRUE)
    expect(viewMain.left).toBe(0)
    expect(viewColumnRight.left).toBe(0)

    // Re-showing is where Univer's zero shift used to leave the grid 46px
    // under the row-header strip.
    handleRibbonCommand(ctx, 'toggle-headings')
    await settle()
    expect(config.rowHeader.hidden).toBe(BooleanNumber.FALSE)
    expect(viewMain.left).toBe(46)
    expect(viewColumnRight.left).toBe(46)
  })
})
