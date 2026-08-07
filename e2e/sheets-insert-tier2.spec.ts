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

function saveWorkbook(app: Awaited<ReturnType<typeof launchShell>>['app']): Promise<void> {
  return app.evaluate(({ webContents }) => {
    const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('sheets/out'))
    wc?.send('menu:action', 'save')
  })
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

test.describe('sheets: Insert → Equation and Checkbox', () => {
  test('renders LaTeX to a picture and saves it into the workbook', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-tier2-e2e-'))
    const workbook = join(scratch, 'equation.xlsx')
    await copyFile(FIXTURE, workbook)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-insert-tier2',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, 'sheets/out')
      await waitForWorkbook(sheets)

      await sheets.getByRole('button', { name: 'Insert', exact: true }).click()
      await sheets.getByRole('button', { name: 'Equation' }).click()
      const dialog = sheets.getByRole('dialog', { name: 'Insert Equation' })
      await expect(dialog).toBeVisible()

      // Presets render as MathML thumbnails; picking one fills the input.
      await expect(dialog.locator('.equation-presets button').first()).toBeVisible()
      await dialog.locator('.equation-latex').fill('x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}')
      await expect(dialog.locator('.equation-preview math')).toBeVisible()
      await sheets.screenshot({ path: screenshotPath('equation-preview') })
      await dialog.getByRole('button', { name: 'OK' }).click()

      await expect(sheets.getByRole('status')).toContainText('Picture inserted', {
        timeout: 20_000,
      })
      await expect(sheets.locator('.xlsx-image')).toBeVisible({ timeout: 15_000 })
      await sheets.screenshot({ path: screenshotPath('equation-inserted') })

      await saveWorkbook(launched.app)
      await expect(() => {
        const listing = execSync(`unzip -l "${workbook}"`).toString()
        expect(listing).toContain('xl/media/')
      }).toPass({ timeout: 20_000 })
    } finally {
      await closeAndSaveVideo(launched, 'sheets-insert-tier2')
    }
  })

  test('inserts a checkbox rule that saves as a two-value list validation', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-tier2-e2e-'))
    const workbook = join(scratch, 'checkbox.xlsx')
    await copyFile(FIXTURE, workbook)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-insert-tier2',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, 'sheets/out')
      await waitForWorkbook(sheets)

      await sheets.getByRole('button', { name: 'Insert', exact: true }).click()
      // DV edits are gated until the sheet's own file rules are installed
      // (indexing complete), so retry until the rule lands in the saved file.
      await expect(async () => {
        await sheets.getByRole('button', { name: 'Checkbox' }).click()
        await saveWorkbook(launched.app)
        const sheetXml = execSync(`unzip -p "${workbook}" xl/worksheets/sheet1.xml`).toString()
        expect(sheetXml).toContain('<dataValidation type="list"')
        expect(sheetXml).toContain('<formula1>"1,0"</formula1>')
      }).toPass({ timeout: 30_000 })
      await sheets.screenshot({ path: screenshotPath('checkbox-inserted') })
    } finally {
      await closeAndSaveVideo(launched, 'sheets-insert-tier2')
    }
  })
})

test.describe('sheets: Insert → Timeline', () => {
  test('filters a pivot by month range through the timeline panel', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-tier2-e2e-'))
    const workbook = join(scratch, 'timeline.xlsx')
    await copyFile(FIXTURE, workbook)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-insert-tier2',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, 'sheets/out')
      await waitForWorkbook(sheets)
      // Full-load mode arrives within a few seconds on this tiny fixture;
      // typing the source table below takes longer than that.
      const origin = await gridOrigin(sheets)

      const table = [
        ['Date', 'Amount'],
        ['2024-01-05', '10'],
        ['2024-02-10', '20'],
        ['2024-03-15', '30'],
        ['2024-01-20', '40'],
      ]
      for (const [rowIndex, row] of table.entries()) {
        for (const [columnIndex, value] of row.entries()) {
          const point = cellPoint(origin, rowIndex, columnIndex)
          await sheets.mouse.click(point.x, point.y)
          await sheets.keyboard.type(value, { delay: 30 })
          await sheets.keyboard.press('Enter')
        }
      }

      const a1 = cellPoint(origin, 0, 0)
      const b5 = cellPoint(origin, 4, 1)
      await sheets.mouse.click(a1.x, a1.y)
      await sheets.keyboard.down('Shift')
      await sheets.mouse.click(b5.x, b5.y)
      await sheets.keyboard.up('Shift')

      await sheets.getByRole('button', { name: 'Insert', exact: true }).click()
      await sheets.getByRole('button', { name: 'PivotTable', exact: true }).click()
      const pivotDialog = sheets.getByRole('dialog', { name: 'Create PivotTable' })
      await expect(pivotDialog).toBeVisible()
      await pivotDialog.locator('input.cell-input').last().fill('E1')
      await sheets.screenshot({ path: screenshotPath('timeline-pivot-dialog') })
      await pivotDialog.getByRole('button', { name: 'Create' }).click()
      await expect(pivotDialog).not.toBeVisible({ timeout: 15_000 })
      await sheets.waitForTimeout(1_500)
      await sheets.screenshot({ path: screenshotPath('timeline-pivot-created') })

      // A freshly created pivot lives only in the edit journal; slicers and
      // timelines bind to file-loaded pivots, and saving reopens the session
      // over the newly written definition.
      await saveWorkbook(launched.app)
      await waitForWorkbook(sheets)

      // Put the cursor inside the pivot output (top-left header cell at E1);
      // retry while the reopened session streams the pivot definition in.
      const picker = sheets.getByRole('dialog', { name: 'Insert Timeline' })
      const e1 = cellPoint(origin, 0, 4)
      await expect(async () => {
        await sheets.mouse.click(e1.x, e1.y)
        await sheets.getByRole('button', { name: 'Insert', exact: true }).click()
        await sheets.getByRole('button', { name: 'Timeline' }).click()
        await expect(picker).toBeVisible({ timeout: 2_000 })
      }).toPass({ timeout: 30_000 })
      await picker.getByRole('button', { name: 'Date' }).click()
      await expect(picker).not.toBeVisible()

      const panel = sheets.locator('.timeline-panel')
      await expect(panel).toBeVisible()
      await expect(panel.locator('.timeline-month')).toHaveCount(3)
      await sheets.screenshot({ path: screenshotPath('timeline-panel') })

      await panel.locator('.timeline-month').first().click()
      await expect(sheets.locator('.workbook-status')).toContainText('applied', {
        timeout: 15_000,
      })
      await expect(panel.locator('.timeline-month.selected')).toHaveCount(1)
      await sheets.screenshot({ path: screenshotPath('timeline-filtered') })

      // Clearing restores all members.
      await panel.locator('.slicer-clear').click()
      await expect(sheets.locator('.workbook-status')).toContainText('cleared', {
        timeout: 15_000,
      })
    } finally {
      await closeAndSaveVideo(launched, 'sheets-insert-tier2')
    }
  })
})
