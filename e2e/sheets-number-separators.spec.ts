import { test, expect } from '@playwright/test'
import { execSync } from 'node:child_process'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

/**
 * genoffice#2883: in a decimal-comma region `5,47` was stored as text, so every
 * formula over it returned #VALUE!. The grid and the editors follow the
 * region's separators like Excel; the saved file stays canonical.
 */
test.describe('sheets: region number separators', () => {
  test('a decimal-comma region types 5,47 and =SUM(A1;A2) and saves canonical values', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-sheets-separators-'))
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-number-separators',
      env: { GENOFFICE_NUMBER_LOCALE: 'pt-BR' },
    })
    try {
      const { app, page } = launched
      await app.evaluate(({ app: electronApp }, dir) => {
        electronApp.setPath('documents', dir)
      }, scratch)
      const workbook = join(scratch, 'separators.xlsx')
      await app.evaluate(({ dialog }, target) => {
        dialog.showSaveDialog = async () => ({ canceled: false, filePath: target })
      }, workbook)

      await page.locator('.quick-card').nth(1).click()
      const sheets = await waitForPageWithUrl(app, '://sheets/')
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
      await sheets.mouse.click(grid.x + 46 + 43, grid.y + 24 + 12)
      const nameBox = sheets.locator('[data-u-comp="defined-name"] input')
      await expect(nameBox).toHaveValue('A1')
      for (const typed of ['5,47', '=A1*1,5', '=SUM(A1;A2)', '1.234,5', '=IF(A1>5,4;"a;b";0,5)']) {
        await sheets.keyboard.type(typed, { delay: 40 })
        await sheets.keyboard.press('Enter')
      }
      await expect(nameBox).toHaveValue('A6')
      await sheets.screenshot({ path: screenshotPath('sheets-number-separators-typed') })

      // re-entering a cell opens it in the region's notation and commits it unchanged
      await sheets.mouse.click(grid.x + 46 + 43, grid.y + 24 + 12 + 20 * 2)
      await expect(nameBox).toHaveValue('A3')
      await sheets.keyboard.press('F2')
      await sheets.screenshot({ path: screenshotPath('sheets-number-separators-editing') })
      await sheets.keyboard.press('Enter')
      // the re-committed formula recalculates before the save reads its value
      await sheets.waitForTimeout(1_000)

      await app.evaluate(({ webContents }) => {
        const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('://sheets/'))
        wc?.send('menu:action', 'save')
      })
      await expect(() => {
        const xml = execSync(`unzip -p "${workbook}" xl/worksheets/sheet1.xml`).toString()
        const cell = (ref: string) =>
          new RegExp(`<c r="${ref}"[^>]*>(.*?)</c>`).exec(xml)?.[1] ?? ''
        expect(cell('A1')).toBe('<v>5.47</v>')
        expect(cell('A2')).toBe('<f>A1*1.5</f><v>8.205</v>')
        expect(cell('A3')).toBe('<f>SUM(A1,A2)</f><v>13.675</v>')
        expect(cell('A4')).toBe('<v>1234.5</v>')
        expect(cell('A5')).toContain('<f>IF(A1&gt;5.4,"a;b",0.5)</f>')
      }).toPass({ timeout: 15_000 })
    } finally {
      await closeAndSaveVideo(launched, 'sheets-number-separators')
    }
  })
})
