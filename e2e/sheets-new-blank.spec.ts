import { test, expect } from '@playwright/test'
import { execSync } from 'node:child_process'
import { mkdtemp, readdir } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

/**
 * Regression for "new spreadsheet cannot be saved" (feedback 2368785): the
 * quick-create card must create a real backing .xlsx up front so the save
 * pipeline works from the first edit.
 */
test.describe('sheets: new blank workbook', () => {
  test('quick-create writes a backing file and saves the first edit', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-sheets-blank-'))
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'sheets-new-blank' })
    try {
      const { app, page } = launched
      // keep the auto-created workbook out of the real ~/Documents/GenOffice
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

      // the backing file exists before any edit
      const saveDir = join(scratch, 'GenOffice')
      const created = (await readdir(saveDir)).filter((f) => f.endsWith('.xlsx'))
      expect(created).toHaveLength(1)
      const workbook = join(saveDir, created[0])

      const grid = await sheets.evaluate(() => {
        for (const canvas of document.querySelectorAll('canvas')) {
          const rect = canvas.getBoundingClientRect()
          if (rect.width > 500 && rect.height > 300) return { x: rect.x, y: rect.y }
        }
        return null
      })
      if (!grid) throw new Error('worksheet canvas not found')
      await sheets.mouse.click(grid.x + 46 + 43, grid.y + 24 + 12)
      await expect(sheets.locator('.name-box')).toHaveValue('A1')
      await sheets.keyboard.type('42', { delay: 50 })
      await sheets.keyboard.press('Enter')
      await sheets.screenshot({ path: screenshotPath('sheets-new-blank-edited') })

      await app.evaluate(({ webContents }) => {
        const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('sheets/out'))
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
