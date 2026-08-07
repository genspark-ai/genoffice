import { test, expect } from '@playwright/test'
import { execSync } from 'node:child_process'
import { copyFile, mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

const FIXTURE = resolve(__dirname, '../apps/sheets/fixtures/generated/compatibility-basic.xlsx')

async function waitForWorkbook(page: Page): Promise<void> {
  await page.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
    timeout: 30_000,
  })
  await page.waitForTimeout(1_500)
}

/** grid origin: the worksheet canvas top-left */
async function gridOrigin(page: Page): Promise<{ x: number; y: number }> {
  const grid = await page.evaluate(() => {
    for (const canvas of document.querySelectorAll('canvas')) {
      const rect = canvas.getBoundingClientRect()
      if (rect.width > 500 && rect.height > 300) return { x: rect.x, y: rect.y }
    }
    return null
  })
  if (!grid) throw new Error('worksheet canvas not found')
  return grid
}

/** center of a cell: ~46px row header, ~24px column header, ~86px × ~23px cells */
function cellPoint(origin: { x: number; y: number }, row: number, column: number) {
  return { x: origin.x + 46 + column * 86 + 43, y: origin.y + 24 + row * 23 + 11 }
}

test.describe('sheets: Insert → Recommended Charts and Icons', () => {
  test('recommends charts for a time series and inserts an icon', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-gallery-e2e-'))
    const workbook = join(scratch, 'gallery.xlsx')
    await copyFile(FIXTURE, workbook)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-insert-gallery',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, 'sheets/out')
      await waitForWorkbook(sheets)
      const origin = await gridOrigin(sheets)

      const table = [
        ['Month', 'Sales'],
        ['Jan', '10'],
        ['Feb', '12'],
        ['Mar', '9'],
      ]
      for (const [rowIndex, row] of table.entries()) {
        for (const [columnIndex, value] of row.entries()) {
          const point = cellPoint(origin, rowIndex, columnIndex)
          await sheets.mouse.click(point.x, point.y)
          await sheets.keyboard.type(value, { delay: 30 })
          await sheets.keyboard.press('Enter')
        }
      }

      // select A1:B4 for the recommendation read
      const a1 = cellPoint(origin, 0, 0)
      const b4 = cellPoint(origin, 3, 1)
      await sheets.mouse.click(a1.x, a1.y)
      await sheets.keyboard.down('Shift')
      await sheets.mouse.click(b4.x, b4.y)
      await sheets.keyboard.up('Shift')

      await sheets.getByRole('button', { name: 'Insert', exact: true }).click()
      await sheets.getByRole('button', { name: 'Recommended Charts' }).click()
      const recoDialog = sheets.getByRole('dialog', { name: 'Recommended Charts' })
      await expect(recoDialog).toBeVisible()
      const recoItems = recoDialog.locator('.recommended-charts-grid button')
      expect(await recoItems.count()).toBeGreaterThanOrEqual(3)
      // month labels → the time-series heuristic must put Line first
      await expect(recoItems.first()).toContainText('Line')
      await sheets.screenshot({ path: screenshotPath('recommended-charts') })
      await recoItems.first().click()
      await expect(recoDialog).not.toBeVisible()
      await expect(sheets.locator('.xlsx-chart')).toBeVisible({ timeout: 15_000 })

      // Icons: search-driven pick rides the picture pipeline
      await sheets.getByRole('button', { name: 'Icons' }).click()
      const iconsDialog = sheets.getByRole('dialog', { name: 'Icons' })
      await expect(iconsDialog).toBeVisible()
      expect(await iconsDialog.locator('.icons-grid button').count()).toBeGreaterThan(10)
      await iconsDialog.getByRole('searchbox').fill('rocket')
      const hits = iconsDialog.locator('.icons-grid button')
      await expect(hits.first()).toBeVisible()
      await sheets.screenshot({ path: screenshotPath('icons-picker') })
      await hits.first().click()
      await expect(iconsDialog).not.toBeVisible({ timeout: 10_000 })
      await expect(sheets.getByRole('status')).toContainText('Picture inserted')
      await expect(sheets.locator('.xlsx-image')).toBeVisible({ timeout: 15_000 })
      await sheets.screenshot({ path: screenshotPath('gallery-inserted') })

      await launched.app.evaluate(({ webContents }) => {
        const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('sheets/out'))
        wc?.send('menu:action', 'save')
      })
      await expect(() => {
        const listing = execSync(`unzip -l "${workbook}"`).toString()
        expect(listing).toContain('xl/media/')
        expect(listing).toContain('xl/charts/')
      }).toPass({ timeout: 20_000 })
    } finally {
      await closeAndSaveVideo(launched, 'sheets-insert-gallery')
    }
  })
})
