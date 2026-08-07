import { test, expect } from '@playwright/test'
import { execSync } from 'node:child_process'
import { copyFile, mkdtemp, writeFile } from 'node:fs/promises'
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

test.describe('sheets: manual visual inserts are undoable', () => {
  test('undo removes an icon inserted from the ribbon', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-polish-e2e-'))
    const workbook = join(scratch, 'undo.xlsx')
    await copyFile(FIXTURE, workbook)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-edit-polish',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, 'sheets/out')
      await waitForWorkbook(sheets)

      await sheets.getByRole('button', { name: 'Insert', exact: true }).click()
      await sheets.getByRole('button', { name: 'Icons' }).click()
      const dialog = sheets.getByRole('dialog', { name: 'Icons' })
      await expect(dialog).toBeVisible()
      await dialog.locator('.icons-grid button').first().click()
      await expect(dialog).not.toBeVisible({ timeout: 10_000 })
      await expect(sheets.locator('.xlsx-image')).toBeVisible({ timeout: 15_000 })

      // Put focus back on the grid so ⌘Z reaches Univer's shortcut handler.
      const grid = await sheets.evaluate(() => {
        for (const canvas of document.querySelectorAll('canvas')) {
          const rect = canvas.getBoundingClientRect()
          if (rect.width > 500 && rect.height > 300) return { x: rect.x, y: rect.y }
        }
        return null
      })
      if (!grid) throw new Error('worksheet canvas not found')
      await sheets.mouse.click(grid.x + 46 + 43, grid.y + 24 + 11 + 230)
      await sheets.keyboard.press('ControlOrMeta+z')
      await expect(sheets.locator('.xlsx-image')).not.toBeVisible({ timeout: 10_000 })
      await sheets.screenshot({ path: screenshotPath('undo-after') })

      await sheets.keyboard.press('ControlOrMeta+y')
      await expect(sheets.locator('.xlsx-image')).toBeVisible({ timeout: 10_000 })
    } finally {
      await closeAndSaveVideo(launched, 'sheets-edit-polish')
    }
  })
})

test.describe('sheets: freeze journal follows undo', () => {
  test('freezing then undoing leaves no pane in the saved file', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-polish-e2e-'))
    const workbook = join(scratch, 'freeze.xlsx')
    await copyFile(FIXTURE, workbook)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-edit-polish',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, 'sheets/out')
      await waitForWorkbook(sheets)

      await sheets.getByRole('button', { name: 'View', exact: true }).click()
      await sheets.getByRole('button', { name: 'Freeze Panes' }).click()
      await sheets.getByRole('option', { name: 'Freeze Top Row' }).click()
      await sheets.waitForTimeout(500)
      // Focus the grid so ⌘Z reaches Univer, then undo the freeze.
      const grid = await sheets.evaluate(() => {
        for (const canvas of document.querySelectorAll('canvas')) {
          const rect = canvas.getBoundingClientRect()
          if (rect.width > 500 && rect.height > 300) return { x: rect.x, y: rect.y }
        }
        return null
      })
      if (!grid) throw new Error('worksheet canvas not found')
      await sheets.mouse.click(grid.x + 46 + 43, grid.y + 24 + 11 + 46)
      await sheets.keyboard.press('ControlOrMeta+z')
      await sheets.waitForTimeout(500)
      // A cell edit keeps the workbook dirty so the save rewrites the sheet.
      await sheets.keyboard.type('marker', { delay: 30 })
      await sheets.keyboard.press('Enter')

      await launched.app.evaluate(({ webContents }) => {
        const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('sheets/out'))
        wc?.send('menu:action', 'save')
      })
      await expect(() => {
        const xml = execSync(`unzip -p "${workbook}" xl/worksheets/sheet1.xml`).toString()
        expect(xml).toContain('marker')
        expect(xml).not.toContain('<pane')
      }).toPass({ timeout: 20_000 })
    } finally {
      await closeAndSaveVideo(launched, 'sheets-edit-polish')
    }
  })
})

test.describe('sheets: Data → From Text/CSV', () => {
  test('imports a CSV file at the selection and saves the values', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-polish-e2e-'))
    const workbook = join(scratch, 'csv.xlsx')
    await copyFile(FIXTURE, workbook)
    const csvPath = join(scratch, 'data.csv')
    await writeFile(csvPath, 'City,Population\nTokyo,37\nParis,11\n', 'utf8')

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-edit-polish',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, 'sheets/out')
      await waitForWorkbook(sheets)

      await sheets.getByRole('button', { name: 'Data', exact: true }).click()
      const chooser = sheets.waitForEvent('filechooser')
      await sheets.getByRole('button', { name: 'From Text/CSV' }).click()
      await (await chooser).setFiles(csvPath)

      await expect(sheets.locator('.workbook-status')).toContainText('A1', { timeout: 15_000 })
      await launched.app.evaluate(({ webContents }) => {
        const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('sheets/out'))
        wc?.send('menu:action', 'save')
      })
      await expect(() => {
        const xml = execSync(`unzip -p "${workbook}" xl/worksheets/sheet1.xml`).toString()
        expect(xml).toContain('Tokyo')
        expect(xml).toContain('<v>37</v>')
      }).toPass({ timeout: 20_000 })
    } finally {
      await closeAndSaveVideo(launched, 'sheets-edit-polish')
    }
  })
})

test.describe('sheets: comment navigation', () => {
  test('previous/next surface the empty-sheet message without notes', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-polish-e2e-'))
    const workbook = join(scratch, 'notes.xlsx')
    await copyFile(FIXTURE, workbook)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-edit-polish',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, 'sheets/out')
      await waitForWorkbook(sheets)

      await sheets.getByRole('button', { name: 'Review', exact: true }).click()
      await sheets.getByRole('button', { name: 'Previous' }).click()
      await expect(sheets.locator('.workbook-status')).toContainText('No comments', {
        timeout: 10_000,
      })
      await sheets.getByRole('button', { name: 'Next' }).click()
      await expect(sheets.locator('.workbook-status')).toContainText('No comments', {
        timeout: 10_000,
      })
    } finally {
      await closeAndSaveVideo(launched, 'sheets-edit-polish')
    }
  })
})
