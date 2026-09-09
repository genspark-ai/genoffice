import { test, expect } from '@playwright/test'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

// the preload exposes window.__genofficeDebug only under this env var
process.env.GENOFFICE_DEBUG_HOOKS = '1'

/**
 * Regression for "typed a count into the right-click insert-N-columns box,
 * pressed Enter, nothing happened" (alpha feedback): Enter in a menu
 * count box must run the row's action like a click, not just commit the
 * number and leave the menu open.
 */
test.describe('sheets: Enter runs the context-menu insert-N action', () => {
  test('insert 3 columns left of B via the count box and Enter', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-menu-enter-'))
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-menu-input-enter',
    })
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

      await sheets.evaluate(async () => {
        const debug = (window as unknown as Record<string, unknown>).__genofficeDebug as {
          univerAPI: {
            getActiveWorkbook(): {
              getActiveSheet(): {
                getRange(
                  row: number,
                  column: number,
                  rows: number,
                  columns: number,
                ): { setValues(v: unknown[][]): Promise<unknown> }
              }
            }
          }
        }
        await debug.univerAPI
          .getActiveWorkbook()
          .getActiveSheet()
          .getRange(0, 1, 1, 1)
          .setValues([['marker']])
      })

      // right-click column B's header → context menu with the count boxes
      await sheets.mouse.click(grid.x + 46 + 86 + 43, grid.y + 12, { button: 'right' })
      const countBox = sheets.locator('input.univer-h-7').first()
      await expect(countBox).toBeVisible()

      await countBox.click()
      await sheets.keyboard.press('ControlOrMeta+a')
      await sheets.keyboard.type('3', { delay: 50 })
      await sheets.keyboard.press('Enter')

      // the insert ran: the marker moved from B1 to E1 and the menu closed
      await expect(async () => {
        const values = await sheets.evaluate(() => {
          const debug = (window as unknown as Record<string, unknown>).__genofficeDebug as {
            univerAPI: {
              getActiveWorkbook(): {
                getActiveSheet(): {
                  getRange(
                    row: number,
                    column: number,
                    rows: number,
                    columns: number,
                  ): { getValues(): unknown[][] }
                }
              }
            }
          }
          return debug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
            .getRange(0, 0, 1, 6)
            .getValues()[0]
        })
        expect(values[4]).toBe('marker')
        expect(values[1]).toBeNull()
      }).toPass({ timeout: 10_000 })
      await expect(countBox).toBeHidden()
    } finally {
      await closeAndSaveVideo(launched, 'sheets-menu-input-enter')
    }
  })
})
