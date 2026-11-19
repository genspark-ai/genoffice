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

async function gridOrigin(page: Page): Promise<{ x: number; y: number }> {
  const grid = await page.evaluate(() => {
    for (const canvas of document.querySelectorAll('canvas')) {
      const rect = canvas.getBoundingClientRect()
      if (rect.width > 500 && rect.height > 300) return { x: rect.x, y: rect.y }
    }
    return null
  })
  if (!grid) throw new Error('worksheet canvas not found')
  return { x: grid.x + 46, y: grid.y + 24 }
}

function sheetXml(workbook: string): string {
  return execSync(`unzip -p "${workbook}" xl/worksheets/sheet1.xml`).toString()
}

function saveWorkbook(app: Awaited<ReturnType<typeof launchShell>>['app']): Promise<void> {
  return app.evaluate(({ webContents }) => {
    const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('://sheets/'))
    wc?.send('menu:action', 'save')
  })
}

async function insertIconPicture(sheets: Page): Promise<void> {
  // The selected tab doubles as the collapse control, so only press it when
  // the band is showing another tab — pressing the active one folds the
  // ribbon away and the Icons button with it.
  const tab = sheets.getByRole('button', { name: 'Insert', exact: true })
  if (!(await tab.evaluate((button) => button.classList.contains('active')))) {
    await tab.click()
  }
  await sheets.getByRole('button', { name: 'Icons' }).click()
  const dialog = sheets.getByRole('dialog', { name: 'Icons' })
  await expect(dialog).toBeVisible()
  await dialog.locator('.icons-grid button').first().click()
  await expect(dialog).not.toBeVisible({ timeout: 10_000 })
  await expect(sheets.locator('.xlsx-image')).toBeVisible({ timeout: 15_000 })
}

test.describe('sheets: picture in cell', () => {
  test('embeds floating pictures, previews them, and round-trips Excel rich values', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-cell-image-e2e-'))
    const workbook = join(scratch, 'cell-image.xlsx')
    await copyFile(FIXTURE, workbook)

    // ── session 1: embed a just-inserted picture, leave a second one in the
    //    file, save ──
    const first = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-cell-image',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(first.app, '://sheets/')
      await waitForWorkbook(sheets)
      const origin = await gridOrigin(sheets)
      const nameBox = sheets.locator('[data-u-comp="defined-name"] input')

      // A floating picture first (Insert → Icons renders one at A1).
      await insertIconPicture(sheets)
      const picture = sheets.locator('.xlsx-image')

      // Right-click → Embed Image in Cell.
      await picture.click({ button: 'right' })
      await sheets.getByRole('button', { name: 'Embed Image in Cell' }).click()

      // The floating picture is gone; the cell shows the picture instead.
      await expect(picture).toHaveCount(0)
      const cellImage = sheets.locator('.xlsx-cell-image')
      await expect(cellImage).toBeVisible({ timeout: 10_000 })
      await sheets.screenshot({ path: screenshotPath('cell-image-embedded') })

      // Double-click opens the zoomable viewer; Escape closes it.
      const box = await cellImage.boundingBox()
      if (!box) throw new Error('cell picture overlay has no box')
      await sheets.mouse.dblclick(box.x + box.width / 2, box.y + box.height / 2)
      const viewer = sheets.getByRole('dialog', { name: 'Image viewer' })
      await expect(viewer).toBeVisible({ timeout: 10_000 })
      await sheets.screenshot({ path: screenshotPath('cell-image-viewer') })
      await sheets.keyboard.press('Escape')
      await expect(viewer).toHaveCount(0)

      // A second picture, anchored to E5, stays a floating file picture.
      await sheets.mouse.click(origin.x + 74 * 4 + 37, origin.y + 20 * 4 + 10)
      await expect(nameBox).toHaveValue('E5')
      await insertIconPicture(sheets)
      await expect(picture).toHaveCount(1)

      // Saving writes Excel's rich-value parts and the vm-marked host cell.
      await saveWorkbook(first.app)
      await expect(() => {
        expect(sheetXml(workbook)).toMatch(/r="A1"[^>]*t="e" vm="1"/)
      }).toPass({ timeout: 20_000 })
      const listing = execSync(`unzip -l "${workbook}"`).toString()
      expect(listing).toContain('xl/metadata.xml')
      expect(listing).toContain('xl/richData/rdrichvalue.xml')
      expect(listing).toContain('xl/richData/rdrichvaluestructure.xml')
      expect(listing).toContain('xl/richData/richValueRel.xml')
      expect(listing).toContain('xl/media/')
      // The second picture is still a drawing in the file.
      expect(listing).toContain('xl/drawings/')
    } finally {
      await closeAndSaveVideo(first, 'sheets-cell-image')
    }

    // ── session 2: reopen; the cell picture renders from the file's rich
    //    data, and the file's floating picture embeds too ──
    const second = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-cell-image-reopen',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(second.app, '://sheets/')
      await waitForWorkbook(sheets)

      const cellImage = sheets.locator('.xlsx-cell-image')
      await expect(cellImage).toBeVisible({ timeout: 15_000 })
      expect(
        await cellImage.locator('img').evaluate((img: HTMLImageElement) => img.naturalWidth),
      ).toBeGreaterThan(0)

      // The file's floating picture embeds through the media IPC.
      const picture = sheets.locator('.xlsx-image')
      await expect(picture).toBeVisible({ timeout: 15_000 })
      await picture.click({ button: 'right' })
      await sheets.getByRole('button', { name: 'Embed Image in Cell' }).click()
      await expect(picture).toHaveCount(0)
      await expect(cellImage).toHaveCount(2)

      await saveWorkbook(second.app)
      await expect(() => {
        expect(sheetXml(workbook)).toMatch(/r="E5"[^>]*t="e" vm="2"/)
      }).toPass({ timeout: 20_000 })
    } finally {
      await closeAndSaveVideo(second, 'sheets-cell-image-reopen')
    }

    // ── session 3: both pictures render from the saved file ──
    const third = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-cell-image-final',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(third.app, '://sheets/')
      await waitForWorkbook(sheets)

      const cellImage = sheets.locator('.xlsx-cell-image')
      await expect(cellImage).toHaveCount(2, { timeout: 15_000 })
      for (const image of await cellImage.locator('img').all()) {
        expect(await image.evaluate((img: HTMLImageElement) => img.naturalWidth)).toBeGreaterThan(0)
      }
      await sheets.screenshot({ path: screenshotPath('cell-image-reopened') })
    } finally {
      await closeAndSaveVideo(third, 'sheets-cell-image-final')
    }
  })
})
