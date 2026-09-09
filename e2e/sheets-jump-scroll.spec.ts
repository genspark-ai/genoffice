import { test, expect } from '@playwright/test'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

// the preload exposes window.__genofficeDebug only under this env var; the
// spec needs it to read Univer's visible range through the Facade
process.env.GENOFFICE_DEBUG_HOOKS = '1'

/**
 * Regression for "Ctrl+Shift+Down extends the selection but the viewport
 * stays put" (alpha feedback): the reveal on SetSelectionsOperation
 * scrolls to the anchor cell, which never moves while extending, so the
 * grid never followed the selection off-screen. The jump-nav override must
 * scroll the moving edge into view — and back when the selection shrinks.
 */

interface GridState {
  visible: { startRow: number; endRow: number }
  selection: { startRow: number; endRow: number }
}

async function gridState(sheets: Page): Promise<GridState> {
  return sheets.evaluate(() => {
    const debug = (window as unknown as Record<string, unknown>).__genofficeDebug as {
      univerAPI: {
        getActiveWorkbook(): {
          getActiveRange(): { getRange(): { startRow: number; endRow: number } } | null
          getActiveSheet(): { getVisibleRange(): { startRow: number; endRow: number } }
        }
      }
    }
    const workbook = debug.univerAPI.getActiveWorkbook()
    const visible = workbook.getActiveSheet().getVisibleRange()
    const selection = workbook.getActiveRange()?.getRange() ?? { startRow: -1, endRow: -1 }
    return {
      visible: { startRow: visible.startRow, endRow: visible.endRow },
      selection: { startRow: selection.startRow, endRow: selection.endRow },
    }
  })
}

test.describe('sheets: ctrl+shift+arrow scroll follow', () => {
  test('extending the selection to the sheet edge scrolls the moving edge into view', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-jump-scroll-'))
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'sheets-jump-scroll' })
    try {
      const { app, page } = launched
      await app.evaluate(({ app: electronApp }, dir) => {
        electronApp.setPath('documents', dir)
      }, scratch)

      await expect(page.locator('.quick-card').nth(1)).toContainText('AI Sheets')
      await page.locator('.quick-card').nth(1).click()

      const sheets = await waitForPageWithUrl(app, 'sheets/out')
      await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
        timeout: 30_000,
      })
      await sheets.waitForTimeout(1_500)

      const grid = await sheets.evaluate(() => {
        for (const canvas of document.querySelectorAll('canvas')) {
          const rect = canvas.getBoundingClientRect()
          if (rect.width > 500 && rect.height > 300) return { x: rect.x, y: rect.y }
        }
        return null
      })
      if (!grid) throw new Error('worksheet canvas not found')

      // a lone entry in A1: Ctrl+Shift+Down extends all the way to the last
      // row. Set the value through the Facade — typing it opens the cell
      // editor, whose hidden focus swallows the arrow shortcut afterwards.
      await sheets.evaluate(async () => {
        const debug = (window as unknown as Record<string, unknown>).__genofficeDebug as {
          univerAPI: {
            getActiveWorkbook(): {
              getActiveSheet(): {
                getRange(row: number, column: number): { setValue(v: number): Promise<unknown> }
              }
            }
          }
        }
        await debug.univerAPI.getActiveWorkbook().getActiveSheet().getRange(0, 0).setValue(42)
      })
      await sheets.mouse.click(grid.x + 46 + 43, grid.y + 24 + 12)
      await expect(sheets.locator('.name-box')).toHaveValue('A1')
      const before = await gridState(sheets)
      expect(before.visible.startRow).toBe(0)
      expect(before.selection).toEqual({ startRow: 0, endRow: 0 })

      await sheets.keyboard.press('Control+Shift+ArrowDown')
      // the selection extends to the last row and the viewport follows the
      // moving bottom edge off the first screen
      await expect(async () => {
        const after = await gridState(sheets)
        expect(after.selection.startRow).toBe(0)
        expect(after.selection.endRow).toBeGreaterThan(before.visible.endRow)
        expect(after.visible.startRow).toBeGreaterThan(before.visible.endRow)
      }).toPass({ timeout: 10_000 })

      // shrinking back to the anchor must scroll the viewport back up
      await sheets.keyboard.press('Control+Shift+ArrowUp')
      await expect(async () => {
        const back = await gridState(sheets)
        expect(back.selection).toEqual({ startRow: 0, endRow: 0 })
        expect(back.visible.startRow).toBe(0)
      }).toPass({ timeout: 10_000 })
    } finally {
      await closeAndSaveVideo(launched, 'sheets-jump-scroll')
    }
  })
})
