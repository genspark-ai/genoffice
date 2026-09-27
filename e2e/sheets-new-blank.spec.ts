import { test, expect } from '@playwright/test'
import { execSync } from 'node:child_process'
import { existsSync } from 'node:fs'
import { mkdtemp, readdir } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

/**
 * Regression for "new spreadsheet cannot be saved" (feedback 2368785), updated
 * for #1036: quick-create no longer drops a file in the default folder. The
 * backing workbook lives in a temp directory, the default folder stays empty,
 * and the first save goes through Save As. The save pipeline must still work
 * from the first edit.
 */
test.describe('sheets: new blank workbook', () => {
  test('quick-create saves the first edit through Save As and leaves the default folder empty', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-sheets-blank-'))
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'sheets-new-blank' })
    try {
      const { app, page } = launched
      // keep the auto-created workbook out of the real ~/Documents/GenOffice
      await app.evaluate(({ app: electronApp }, dir) => {
        electronApp.setPath('documents', dir)
      }, scratch)
      const saveDir = join(scratch, 'GenOffice')
      const workbook = join(saveDir, 'quick-create.xlsx')
      // the workbook has no file yet, so Save answers the Save As picker
      await app.evaluate(({ dialog }, target) => {
        dialog.showSaveDialog = async () => ({ canceled: false, filePath: target })
      }, workbook)

      await expect(page.locator('.quick-card').nth(1)).toContainText('AI Sheets')
      await page.locator('.quick-card').nth(1).click()

      const sheets = await waitForPageWithUrl(app, '://sheets/')
      await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
        timeout: 30_000,
      })
      await sheets.waitForTimeout(1_500)

      // nothing lands in the default folder before the user saves (#1036)
      const before = existsSync(saveDir)
        ? (await readdir(saveDir)).filter((f) => f.endsWith('.xlsx'))
        : []
      expect(before).toHaveLength(0)

      const grid = await sheets.evaluate(() => {
        for (const canvas of document.querySelectorAll('canvas')) {
          const rect = canvas.getBoundingClientRect()
          if (rect.width > 500 && rect.height > 300) return { x: rect.x, y: rect.y }
        }
        return null
      })
      if (!grid) throw new Error('worksheet canvas not found')
      await sheets.mouse.click(grid.x + 46 + 43, grid.y + 24 + 12)
      await expect(sheets.locator('[data-u-comp="defined-name"] input')).toHaveValue('A1')
      await sheets.keyboard.type('42', { delay: 50 })
      await sheets.keyboard.press('Enter')
      await sheets.screenshot({ path: screenshotPath('sheets-new-blank-edited') })

      await app.evaluate(({ webContents }) => {
        const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('://sheets/'))
        wc?.send('menu:action', 'save')
      })
      await expect(() => {
        const xml = execSync(`unzip -p "${workbook}" xl/worksheets/sheet1.xml`).toString()
        expect(xml).toContain('<v>42</v>')
      }).toPass({ timeout: 15_000 })
    } finally {
      await closeAndSaveVideo(launched, 'sheets-new-blank')
    }
  })
})
