import { test, expect } from '@playwright/test'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

// the preload exposes window.__genofficeDebug only under this env var; the
// spec seeds cells and reads them back through Univer's Facade
process.env.GENOFFICE_DEBUG_HOOKS = '1'

/**
 * Excel's Ctrl+drag on the fill handle (genoffice#808): a lone number fills
 * a series instead of repeating, a two-cell series repeats instead of
 * extending. Plain drags keep Univer's defaults.
 */

// Chromium on macOS turns Ctrl+click into a context menu; Cmd is the Mac key
const MODIFIER = process.platform === 'darwin' ? 'Meta' : 'Control'

type Rect = { x: number; y: number; width: number; height: number }

interface FacadeRange {
  setValues(values: number[][]): unknown
  activate(): unknown
  getCellRect(): Rect
  getValues(): unknown[][]
}

function range(page: Page, row: number, column: number, rows = 1, columns = 1) {
  return page.evaluateHandle(
    ([r, c, nr, nc]) => {
      const debug = (window as unknown as Record<string, unknown>).__genofficeDebug as {
        univerAPI: {
          getActiveWorkbook(): {
            getActiveSheet(): {
              getRange(row: number, column: number, rows: number, columns: number): FacadeRange
            }
          }
        }
      }
      return debug.univerAPI.getActiveWorkbook().getActiveSheet().getRange(r, c, nr, nc)
    },
    [row, column, rows, columns] as const,
  )
}

async function seed(page: Page, row: number, column: number, values: number[][]) {
  const cells = await range(page, row, column, values.length, values[0].length)
  await cells.evaluate((r, v) => {
    r.setValues(v)
    r.activate()
  }, values)
}

async function columnValues(page: Page, column: number, rows: number): Promise<unknown[]> {
  const cells = await range(page, 0, column, rows, 1)
  return cells.evaluate((r) => r.getValues().map((row) => row[0]))
}

/** drag the active selection's fill handle down so the last filled row is `toRow` */
async function dragFillDown(
  page: Page,
  grid: { x: number; y: number },
  active: { row: number; column: number; rows: number },
  toRow: number,
  modifier?: string,
) {
  const first = await (
    await range(page, active.row, active.column)
  ).evaluate((r) => r.getCellRect())
  const last = await (
    await range(page, active.row + active.rows - 1, active.column)
  ).evaluate((r) => r.getCellRect())
  const handle = { x: grid.x + last.x + last.width, y: grid.y + last.y + last.height }
  const rowPitch = first.height
  const target = {
    x: handle.x - 4,
    y: grid.y + last.y + last.height + (toRow - active.row - active.rows) * rowPitch + rowPitch / 2,
  }
  if (modifier) await page.keyboard.down(modifier)
  await page.mouse.move(handle.x, handle.y)
  await page.mouse.down()
  await page.mouse.move(handle.x, handle.y + 8, { steps: 3 })
  await page.mouse.move(target.x, target.y, { steps: 12 })
  await page.mouse.up()
  if (modifier) await page.keyboard.up(modifier)
}

test.describe('sheets: modifier-drag on the fill handle flips copy and series', () => {
  test('a lone number repeats on a plain drag and counts up under the modifier', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-ctrl-fill-'))
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'sheets-ctrl-drag-fill' })
    try {
      const { app, page } = launched
      await app.evaluate(({ app: electronApp }, dir) => {
        electronApp.setPath('documents', dir)
      }, scratch)

      await expect(page.locator('.quick-card').nth(1)).toContainText('AI Sheets')
      await page.locator('.quick-card').nth(1).click()

      const sheets = await waitForPageWithUrl(app, '://sheets/')
      await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
        timeout: 30_000,
      })
      const gridHandle = await sheets.waitForFunction(
        () => {
          for (const canvas of document.querySelectorAll('canvas')) {
            const rect = canvas.getBoundingClientRect()
            if (rect.width > 500 && rect.height > 300) return { x: rect.x, y: rect.y }
          }
          return null
        },
        null,
        { timeout: 30_000 },
      )
      const grid = (await gridHandle.jsonValue()) as { x: number; y: number }
      await sheets.waitForTimeout(1_000)
      // focus the grid before the Facade activates a cell
      await sheets.mouse.click(grid.x + 46 + 43, grid.y + 24 + 12)

      // A1 = 1008, plain drag to A4: Univer (like Excel) repeats the value.
      // A fresh workbook still streams in for a moment and auto-fill into
      // rows that have not landed is refused, so the first drag retries.
      await expect(async () => {
        await seed(sheets, 0, 0, [[1008]])
        await dragFillDown(sheets, grid, { row: 0, column: 0, rows: 1 }, 3)
        await sheets.waitForTimeout(300)
        expect(await columnValues(sheets, 0, 4)).toEqual([1008, 1008, 1008, 1008])
      }).toPass({ timeout: 30_000, intervals: [1_000] })

      // C1 = 1008, modifier drag to C4: a series
      await seed(sheets, 0, 2, [[1008]])
      await dragFillDown(sheets, grid, { row: 0, column: 2, rows: 1 }, 3, MODIFIER)
      await expect(async () => {
        expect(await columnValues(sheets, 2, 4)).toEqual([1008, 1009, 1010, 1011])
      }).toPass({ timeout: 10_000 })

      // E1:E2 = 1, 2, modifier drag to E4: repeats instead of extending
      await seed(sheets, 0, 4, [[1], [2]])
      await dragFillDown(sheets, grid, { row: 0, column: 4, rows: 2 }, 3, MODIFIER)
      await expect(async () => {
        expect(await columnValues(sheets, 4, 4)).toEqual([1, 2, 1, 2])
      }).toPass({ timeout: 10_000 })
    } finally {
      await closeAndSaveVideo(launched, 'sheets-ctrl-drag-fill')
    }
  })
})
