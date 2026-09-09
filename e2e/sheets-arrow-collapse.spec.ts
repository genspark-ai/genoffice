import { test, expect } from '@playwright/test'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

// the preload exposes window.__genofficeDebug only under this env var; the
// spec needs it to read the selection through Univer's Facade
process.env.GENOFFICE_DEBUG_HOOKS = '1'

/**
 * Regression for "cut K97:L97, press Right, lands on M97" (alpha
 * feedback): an arrow on a multi-cell selection must collapse to the ACTIVE
 * cell and move one step from it (Excel), not step past the range's edge.
 */

async function selectionRect(sheets: Page): Promise<{
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}> {
  return sheets.evaluate(() => {
    const debug = (window as unknown as Record<string, unknown>).__genofficeDebug as {
      univerAPI: {
        getActiveWorkbook(): {
          getActiveRange(): {
            getRange(): {
              startRow: number
              endRow: number
              startColumn: number
              endColumn: number
            }
          } | null
        }
      }
    }
    const range = debug.univerAPI.getActiveWorkbook().getActiveRange()?.getRange()
    if (!range) throw new Error('no active range')
    const { startRow, endRow, startColumn, endColumn } = range
    return { startRow, endRow, startColumn, endColumn }
  })
}

test.describe('sheets: arrow collapses a multi-cell selection to the active cell', () => {
  test('after selecting two cells and cutting, ArrowRight lands next to the active cell', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-arrow-collapse-'))
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-arrow-collapse',
    })
    try {
      const { app, page } = launched
      await app.evaluate(({ app: electronApp }, dir) => {
        electronApp.setPath('documents', dir)
      }, scratch)

      await expect(page.locator('.quick-card').nth(1)).toContainText('AI Sheets')
      await page.locator('.quick-card').nth(1).click()

      const sheets = await waitForPageWithUrl(app, 'sheets/out')
      // the replayed move must not surface as an uncaught command error
      const commandErrors: string[] = []
      sheets.on('pageerror', (err) => {
        if (err.message.includes('[CommandService]')) commandErrors.push(err.message)
      })
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

      // focus the grid, then activate B2 through the Facade (canvas hit
      // coordinates depend on default row/column sizing)
      await sheets.mouse.click(grid.x + 46 + 43, grid.y + 24 + 12)
      await sheets.evaluate(() => {
        const debug = (window as unknown as Record<string, unknown>).__genofficeDebug as {
          univerAPI: {
            getActiveWorkbook(): {
              getActiveSheet(): {
                getRange(
                  row: number,
                  column: number,
                  rows: number,
                  columns: number,
                ): {
                  activate(): unknown
                }
              }
            }
          }
        }
        debug.univerAPI.getActiveWorkbook().getActiveSheet().getRange(1, 1, 1, 1).activate()
      })
      // B2, extend right to C2 (active stays B2), cut — the reported flow
      await sheets.keyboard.press('Shift+ArrowRight')
      expect(await selectionRect(sheets)).toEqual({
        startRow: 1,
        endRow: 1,
        startColumn: 1,
        endColumn: 2,
      })
      await sheets.keyboard.press('Control+x')

      // ArrowRight: collapse to active B2, step one right → C2 (not D2)
      await sheets.keyboard.press('ArrowRight')
      await expect(async () => {
        expect(await selectionRect(sheets)).toEqual({
          startRow: 1,
          endRow: 1,
          startColumn: 2,
          endColumn: 2,
        })
      }).toPass({ timeout: 10_000 })

      // and with the anchor on the far end: select C2:B2 leftwards from C2,
      // active C2 — ArrowDown lands on C3
      await sheets.keyboard.press('Shift+ArrowLeft')
      await sheets.keyboard.press('ArrowDown')
      await expect(async () => {
        expect(await selectionRect(sheets)).toEqual({
          startRow: 2,
          endRow: 2,
          startColumn: 2,
          endColumn: 2,
        })
      }).toPass({ timeout: 10_000 })

      // Ctrl+Arrow (excel-jump) on a multi-cell selection collapses the same
      // way: C3:D3 with active C3, Ctrl+Right jumps along row 3 from C3
      await sheets.keyboard.press('Shift+ArrowRight')
      await sheets.keyboard.press('Control+ArrowRight')
      await expect(async () => {
        const rect = await selectionRect(sheets)
        expect(rect.startRow).toBe(2)
        expect(rect.endRow).toBe(2)
        expect(rect.startColumn).toBe(rect.endColumn)
        expect(rect.startColumn).toBeGreaterThan(2)
      }).toPass({ timeout: 10_000 })
      expect(commandErrors).toEqual([])
    } finally {
      await closeAndSaveVideo(launched, 'sheets-arrow-collapse')
    }
  })
})
