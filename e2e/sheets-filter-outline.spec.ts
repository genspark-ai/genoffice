import { test, expect } from '@playwright/test'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

// the preload exposes window.__genofficeDebug only under this env var
process.env.GENOFFICE_DEBUG_HOOKS = '1'

interface CellRect {
  left: number
  top: number
  right: number
  bottom: number
}

function cellRect(page: Page, row: number, column: number): Promise<CellRect> {
  return page.evaluate(
    ([r, c]) => {
      const debug = (window as unknown as Record<string, unknown>).__genofficeDebug as {
        univerAPI: { getActiveWorkbook(): { getActiveSheet(): unknown } }
      }
      const sheet = debug.univerAPI.getActiveWorkbook().getActiveSheet() as {
        getRange(row: number, column: number): { getCellRect(): CellRect }
      }
      const { left, top, right, bottom } = sheet.getRange(r, c).getCellRect()
      return { left, top, right, bottom }
    },
    [row, column],
  )
}

/**
 * Largest channel spread over a band of the worksheet canvas (sheet coords):
 * gridlines, text and the page are grey, a SelectionControl border is not.
 */
function maxSpread(page: Page, x: number, y: number, w: number, h: number): Promise<number> {
  return page.evaluate(
    ([x, y, w, h]) => {
      let grid: HTMLCanvasElement | null = null
      for (const canvas of document.querySelectorAll('canvas')) {
        const rect = canvas.getBoundingClientRect()
        if (rect.width > 500 && rect.height > 300) grid = canvas
      }
      if (!grid) throw new Error('worksheet canvas not found')
      const dpr = grid.width / grid.getBoundingClientRect().width
      const ctx = grid.getContext('2d')
      if (!ctx) throw new Error('worksheet canvas has no 2d context')
      const data = ctx.getImageData(
        Math.round(x * dpr),
        Math.round(y * dpr),
        Math.max(1, Math.round(w * dpr)),
        Math.max(1, Math.round(h * dpr)),
      ).data
      let spread = 0
      for (let i = 0; i < data.length; i += 4) {
        const r = data[i]
        const g = data[i + 1]
        const b = data[i + 2]
        spread = Math.max(spread, Math.max(r, g, b) - Math.min(r, g, b))
      }
      return spread
    },
    [x, y, w, h],
  )
}

const BAND = 2

function bottomEdge(page: Page, rect: CellRect): Promise<number> {
  const inset = 6
  return maxSpread(
    page,
    rect.left + inset,
    rect.bottom - BAND,
    rect.right - rect.left - 2 * inset,
    2 * BAND + 1,
  )
}

function rightEdge(page: Page, rect: CellRect): Promise<number> {
  const inset = 6
  return maxSpread(
    page,
    rect.right - BAND,
    rect.top + inset,
    2 * BAND + 1,
    rect.bottom - rect.top - 2 * inset,
  )
}

/**
 * Regression for "a mysterious outer border appeared" (alpha
 * feedback): Univer's filter render controller paints a selection-style border
 * around the whole filter range whenever a sheet has a filter. Excel draws
 * no such outline; the range painter is stubbed out at render-module
 * registration (see filter-range-outline.ts). The pixel probe asserts the
 * border colour is gone from the edges the outline used to cover, and the
 * same probe on the parked active cell proves it does see a selection border.
 */
test.describe('sheets: no outline around a filtered range', () => {
  test('creating a filter draws funnel buttons but no range border', async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-filter-outline',
    })
    try {
      const { app, page } = launched

      await expect(page.locator('.quick-card').nth(1)).toContainText('AI Sheets')
      await page.locator('.quick-card').nth(1).click()

      const sheets = await waitForPageWithUrl(app, 'sheets/out')
      await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
        timeout: 30_000,
      })
      await sheets.waitForTimeout(1_500)

      await sheets.evaluate(async () => {
        const debug = (window as unknown as Record<string, unknown>).__genofficeDebug as {
          univerAPI: { getActiveWorkbook(): { getActiveSheet(): unknown } }
        }
        const sheet = debug.univerAPI.getActiveWorkbook().getActiveSheet() as {
          getRange(
            row: number,
            column: number,
            rows: number,
            columns: number,
          ): { setValues(values: unknown[][]): Promise<unknown>; createFilter(): unknown }
        }
        await sheet.getRange(0, 0, 3, 2).setValues([
          ['Name', 'Cat'],
          ['a', 'keep'],
          ['b', 'drop'],
        ])
        sheet.getRange(0, 0, 3, 2).createFilter()
      })
      await sheets.waitForTimeout(1_000)

      // deselect: park the selection far from the filter range so its own
      // selection border cannot be mistaken for the filter outline
      await sheets.evaluate(() => {
        const debug = (window as unknown as Record<string, unknown>).__genofficeDebug as {
          univerAPI: { getActiveWorkbook(): { getActiveSheet(): unknown } }
        }
        const sheet = debug.univerAPI.getActiveWorkbook().getActiveSheet() as {
          getRange(
            row: number,
            column: number,
            rows: number,
            columns: number,
          ): { activate(): unknown }
        }
        sheet.getRange(19, 7, 1, 1).activate()
      })
      await sheets.waitForTimeout(500)

      await sheets.screenshot({ path: screenshotPath('sheets-filter-outline') })

      const hasFilter = await sheets.evaluate(() => {
        const debug = (window as unknown as Record<string, unknown>).__genofficeDebug as {
          univerAPI: { getActiveWorkbook(): { getActiveSheet(): unknown } }
        }
        const sheet = debug.univerAPI.getActiveWorkbook().getActiveSheet() as {
          getFilter(): unknown
        }
        return sheet.getFilter() != null
      })
      expect(hasFilter).toBe(true)

      // positive control: the parked active cell's own selection border is
      // what the filter outline looked like, and the probe must see it
      const parked = await cellRect(sheets, 19, 7)
      expect(await bottomEdge(sheets, parked)).toBeGreaterThan(80)

      const topLeft = await cellRect(sheets, 0, 0)
      const bottomRight = await cellRect(sheets, 2, 1)
      const range = {
        left: topLeft.left,
        top: topLeft.top,
        right: bottomRight.right,
        bottom: bottomRight.bottom,
      }
      expect(await bottomEdge(sheets, range)).toBeLessThan(40)
      expect(await rightEdge(sheets, range)).toBeLessThan(40)
    } finally {
      await closeAndSaveVideo(launched, 'sheets-filter-outline')
    }
  })
})
