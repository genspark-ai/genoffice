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

/** center of cell A1: right of the ~46px row header, below the ~24px column header */
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

function sheetXml(workbookPath: string): string {
  return execSync(`unzip -p "${workbookPath}" xl/worksheets/sheet1.xml`).toString()
}

test.describe('sheets: ribbon batch-1 features', () => {
  test('headings toggle, zoom to selection, names, watch window, calc options', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-ribbon-e2e-'))
    const workbook = join(scratch, 'ribbon-batch.xlsx')
    await copyFile(FIXTURE, workbook)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-ribbon-batch',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, 'sheets/out')
      await waitForWorkbook(sheets)
      const status = sheets.locator('.workbook-status')

      // ── Formulas > Create from Selection: drag A1:B3, create names from top row ──
      const origin = await gridOrigin(sheets)
      await sheets.mouse.move(origin.x + 43, origin.y + 12)
      await sheets.mouse.down()
      await sheets.mouse.move(origin.x + 87 + 43, origin.y + 24 * 2 + 12, { steps: 8 })
      await sheets.mouse.up()
      await sheets.getByRole('button', { name: 'Formulas' }).click()
      await sheets.locator('span.styles-row', { hasText: 'Create from Selection' }).click()
      await sheets.getByRole('option', { name: 'From top row' }).click()
      await expect(status).toContainText('Created 2 names', { timeout: 5_000 })

      // ── Use in Formula lists the new names ──
      await sheets.locator('span.styles-row', { hasText: 'Use in Formula' }).click()
      await expect(sheets.getByRole('listbox')).toBeVisible()
      await sheets.keyboard.press('Escape')

      // ── Watch Window: open, add the selection, a row appears ──
      await sheets.getByRole('button', { name: 'Watch Window' }).click()
      const watch = sheets.locator('.watch-window')
      await expect(watch).toBeVisible()
      await watch.getByRole('button', { name: 'Add watch' }).click()
      await expect(watch.locator('tbody tr')).toHaveCount(6)
      await sheets.screenshot({ path: screenshotPath('sheets-watch-window') })

      // ── Calculation Options: switch to manual, then Calculate Now ──
      await sheets.locator('.ribbon-tool', { hasText: 'Calculation Options' }).click()
      await sheets.getByRole('option', { name: /Manual/ }).click()
      await expect(status).toContainText('Manual calculation on')
      await sheets.locator('button.styles-row', { hasText: 'Calculate Now' }).click()
      await expect(status).toContainText('Recalculated.')
      await sheets.locator('.ribbon-tool', { hasText: 'Calculation Options' }).click()
      await sheets.getByRole('option', { name: /Automatic/ }).click()

      // ── View > Zoom to Selection reports the new zoom ──
      await sheets.getByRole('button', { name: 'View', exact: true }).click()
      await sheets.getByRole('button', { name: 'Zoom to Selection' }).click()
      await expect(status).toContainText('%')

      // ── View > Headings toggle persists as showRowColHeaders="0" on save ──
      await sheets.locator('button.check-item', { hasText: 'Headings' }).click()
      await expect(status).toContainText('Headings hidden.')
      await sheets.screenshot({ path: screenshotPath('sheets-headings-hidden') })
      await launched.app.evaluate(({ webContents }) => {
        const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('sheets/out'))
        wc?.send('menu:action', 'save')
      })
      await expect(() => {
        expect(sheetXml(workbook)).toContain('showRowColHeaders="0"')
      }).toPass({ timeout: 15_000 })
    } finally {
      await closeAndSaveVideo(launched, 'sheets-ribbon-batch')
    }

    // ── reopen: headers stay hidden (loads through the snapshot path) ──
    const second = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-ribbon-reopen',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(second.app, 'sheets/out')
      await waitForWorkbook(sheets)
      await sheets.getByRole('button', { name: 'View', exact: true }).click()
      const headingsBox = sheets
        .locator('button.check-item', { hasText: 'Headings' })
        .locator('.check-box')
      await expect(headingsBox).toHaveText('')
      await sheets.screenshot({ path: screenshotPath('sheets-headings-reopen') })
    } finally {
      await closeAndSaveVideo(second, 'sheets-ribbon-reopen')
    }
  })
})

test.describe('sheets: ribbon batch-2 features', () => {
  test('error checking, goal seek, refresh all', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-ribbon2-e2e-'))
    const workbook = join(scratch, 'ribbon-batch2.xlsx')
    await copyFile(FIXTURE, workbook)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-ribbon-batch2',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, 'sheets/out')
      await waitForWorkbook(sheets)
      const status = sheets.locator('.workbook-status')
      const origin = await gridOrigin(sheets)
      const cell = (column: number, row: number): { x: number; y: number } => ({
        x: origin.x + 87 * column + 43,
        y: origin.y + 24 * row + 12,
      })

      // ── seed: D1 = 1/0 (error), D2 = D3*2 (goal-seek target) ──
      await sheets.mouse.click(cell(3, 0).x, cell(3, 0).y)
      await sheets.keyboard.type('=1/0', { delay: 30 })
      await sheets.keyboard.press('Enter')
      await sheets.keyboard.type('=D3*2', { delay: 30 })
      await sheets.keyboard.press('Enter')
      await sheets.waitForTimeout(500)

      // ── Formulas > Error Checking finds and selects D1 ──
      await sheets.getByRole('button', { name: 'Formulas' }).click()
      await sheets.getByRole('button', { name: 'Error Checking' }).click()
      await expect(status).toContainText('1 errors — at D1: #DIV/0!')
      await expect(sheets.locator('.name-box')).toHaveValue('D1')

      // ── Data > What-If > Goal Seek: D2 = 40 by changing D3 ──
      await sheets.getByRole('button', { name: 'Data', exact: true }).click()
      await sheets.locator('.ribbon-tool', { hasText: 'What-If Analysis' }).click()
      await sheets.getByRole('option', { name: 'Goal Seek' }).click()
      const dialog = sheets.getByRole('dialog', { name: 'Goal Seek' })
      await dialog.getByPlaceholder('B5').fill('D2')
      await dialog.getByPlaceholder('0').fill('40')
      await dialog.getByPlaceholder('B2').fill('D3')
      await dialog.getByRole('button', { name: 'Solve' }).click()
      await expect(dialog.getByText(/Found a solution/)).toBeVisible({ timeout: 30_000 })
      await sheets.screenshot({ path: screenshotPath('sheets-goal-seek') })
      await dialog.getByRole('button', { name: 'Close' }).click()

      // ── Data > Refresh All on a pivot-less workbook reports cleanly ──
      await sheets.locator('button.styles-row', { hasText: 'Refresh All' }).click()
      await expect(status).toContainText('No pivot tables')
    } finally {
      await closeAndSaveVideo(launched, 'sheets-ribbon-batch2')
    }
  })
})

const MINIMAL_THEME =
  '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">' +
  '<a:themeElements><a:clrScheme name="Office">' +
  '<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
  '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
  '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
  '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
  '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>' +
  '</a:clrScheme><a:fontScheme name="Office">' +
  '<a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>' +
  '<a:minorFont><a:latin typeface="Calibri"/></a:minorFont>' +
  '</a:fontScheme><a:fmtScheme name="Office"/></a:themeElements></a:theme>'

function archiveEntry(workbookPath: string, entry: string): string {
  return execSync(`unzip -p "${workbookPath}" ${entry}`).toString()
}

test.describe('sheets: ribbon batch-3 features', () => {
  test('page breaks + preview, workbook protection, allow edit ranges, theme', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-ribbon3-e2e-'))
    const workbook = join(scratch, 'ribbon-batch3.xlsx')
    await copyFile(FIXTURE, workbook)
    // The generated fixture ships no theme part; the theme engine needs one.
    execSync(
      `mkdir -p "${scratch}/xl/theme" && printf '%s' '${MINIMAL_THEME}' > "${scratch}/xl/theme/theme1.xml" ` +
        `&& cd "${scratch}" && zip -q "${workbook}" xl/theme/theme1.xml`,
    )

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-ribbon-batch3',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, 'sheets/out')
      await waitForWorkbook(sheets)
      const status = sheets.locator('.workbook-status')
      const origin = await gridOrigin(sheets)

      // ── Page Layout > Breaks: insert a break at B3 ──
      await sheets.mouse.click(origin.x + 87 + 43, origin.y + 24 * 2 + 12)
      await sheets.getByRole('button', { name: 'Page Layout' }).click()
      await sheets.locator('.ribbon-tool', { hasText: 'Breaks' }).click()
      await sheets.getByRole('option', { name: 'Insert Page Break' }).click()
      await expect(status).toContainText('page break inserted')

      // ── View > Page Break Preview: watermark overlay appears ──
      await sheets.getByRole('button', { name: 'View', exact: true }).click()
      await sheets.getByRole('button', { name: 'Page Break Preview' }).click()
      await expect(status).toContainText('Page Break Preview on.')
      await expect(sheets.locator('.page-break-watermark').first()).toBeVisible()
      await sheets.screenshot({ path: screenshotPath('sheets-page-break-preview') })

      // ── View > Formula Bar toggle hides Univer's fx row and comes back ──
      const fxBar = sheets.locator('[data-u-comp="formula-bar"]')
      await expect(fxBar).toBeVisible()
      await sheets.locator('button.check-item', { hasText: 'Formula Bar' }).click()
      await expect(status).toContainText('Formula bar hidden.')
      await expect(fxBar).toBeHidden()
      await sheets.locator('button.check-item', { hasText: 'Formula Bar' }).click()
      await expect(fxBar).toBeVisible()

      // ── Review > Protect Workbook toggles the structure lock ──
      await sheets.getByRole('button', { name: 'Review', exact: true }).click()
      await sheets.getByRole('button', { name: 'Protect Workbook' }).click()
      await expect(status).toContainText('Workbook structure protection will be written')

      // ── Review > Allow Edit Ranges: add one range ──
      await sheets.getByRole('button', { name: 'Allow Edit Ranges' }).click()
      const dialog = sheets.getByRole('dialog', { name: 'Allow Edit Ranges' })
      await dialog.getByPlaceholder('A1:B4').fill('A1:B2')
      await dialog.getByRole('button', { name: 'Add' }).click()
      await sheets.screenshot({ path: screenshotPath('sheets-allow-edit-ranges') })
      await dialog.getByRole('button', { name: 'OK' }).click()
      await expect(status).toContainText('1 allow-edit range')

      // ── Page Layout > Themes: apply Indigo (grid recolors live) ──
      await sheets.getByRole('button', { name: 'Page Layout' }).click()
      await sheets.locator('.ribbon-tool', { hasText: 'Themes' }).click()
      await sheets.getByRole('option', { name: 'Indigo' }).click()
      await expect(status).toContainText('Theme "Indigo" applied')

      // ── save, then verify every feature landed in the package ──
      await launched.app.evaluate(({ webContents }) => {
        const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('sheets/out'))
        wc?.send('menu:action', 'save')
      })
      await expect(() => {
        const sheet = sheetXml(workbook)
        expect(sheet).toContain('<rowBreaks count="1" manualBreakCount="1">')
        expect(sheet).toContain('<brk id="2" max="16383" man="1"/>')
        expect(sheet).toContain('<brk id="1" max="1048575" man="1"/>')
        expect(sheet).toContain('<protectedRange sqref="A1:B2" name="Range1"/>')
        expect(archiveEntry(workbook, 'xl/workbook.xml')).toContain('lockStructure="1"')
        const theme = archiveEntry(workbook, 'xl/theme/theme1.xml')
        expect(theme).toContain('<a:clrScheme name="Indigo">')
        expect(theme).toContain('<a:accent1><a:srgbClr val="2E4FA3"/></a:accent1>')
        expect(theme).toContain('<a:latin typeface="Segoe UI"/>')
      }).toPass({ timeout: 15_000 })
    } finally {
      await closeAndSaveVideo(launched, 'sheets-ribbon-batch3')
    }
  })
})
