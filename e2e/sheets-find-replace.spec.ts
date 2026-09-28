import { test, expect } from '@playwright/test'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import type { Page } from '@playwright/test'
import { launchShell, waitForPageWithUrl } from './helpers'

/** Minimal single-sheet workbook with three "apple" hits in column A. */
async function buildWorkbook(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="Fruit" sheetId="1" r:id="rId1"/></sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '</Relationships>',
  )
  const rows = ['Item', 'apple pie', 'banana', 'Apple juice', 'cherry', 'apple']
    .map(
      (value, index) =>
        `<row r="${index + 1}"><c r="A${index + 1}" t="inlineStr"><is><t>${value}</t></is></c></row>`,
    )
    .join('')
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      // the dimension keeps the streamed import's row accounting exact
      `<dimension ref="A1:A6"/><sheetData>${rows}</sheetData></worksheet>`,
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

async function waitForWorkbook(page: Page): Promise<void> {
  await page.waitForFunction(() => document.body.textContent?.includes('Fruit'), undefined, {
    timeout: 30_000,
  })
  await page.waitForFunction(() => document.body.textContent?.includes('fully loaded'), undefined, {
    timeout: 60_000,
  })
  await page.waitForTimeout(1_000)
}

// Ctrl+F opens the app's Excel-style Find & Replace panel (the stock Univer
// dialog stays hidden). The panel searches as you type, lists every match
// under Find All, replaces across the session, and closes on Escape.
test('sheets: Excel-style Find & Replace panel', async () => {
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-findreplace-e2e-'))
  const workbook = join(scratch, 'find-replace.xlsx')
  await writeFile(workbook, await buildWorkbook())

  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'sheets-find-replace',
    openFile: workbook,
    lang: 'en',
  })
  try {
    const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
    await waitForWorkbook(sheets)

    const grid = await sheets.locator('canvas').first().boundingBox()
    if (!grid) throw new Error('no grid canvas')
    // click well inside the cell area — a header click does not focus the unit
    await sheets.mouse.click(grid.x + 300, grid.y + 150)
    await sheets.waitForTimeout(400)

    await sheets.keyboard.press('Control+f')
    const panel = sheets.locator('.find-replace-panel')
    await expect(panel).toBeVisible({ timeout: 10_000 })
    // the stock Univer dialog must stay invisible
    const stock = sheets.locator('[data-u-comp="find-replace-dialog"]')
    if ((await stock.count()) > 0) await expect(stock).toBeHidden()

    // search-as-you-type populates the match counter
    await panel.locator('input').first().fill('apple')
    await expect(panel.locator('.fr-status')).toContainText('3', { timeout: 10_000 })

    // Find All lists every match with sheet + address + content
    await panel.getByRole('button', { name: 'Find All' }).click()
    const rows = panel.locator('.fr-results tbody tr')
    await expect(rows).toHaveCount(3, { timeout: 10_000 })
    await expect(rows.first()).toContainText('A2')
    await expect(rows.first()).toContainText('apple pie')

    // Replace tab: replace all occurrences
    await panel.getByRole('tab', { name: 'Replace' }).click()
    const inputs = panel.locator('.fr-grid input:not([type="checkbox"])')
    await inputs.nth(1).fill('pear')
    await sheets.waitForTimeout(1_200)
    await panel.getByRole('button', { name: 'Replace All' }).click()
    await expect(panel.locator('.fr-status')).toContainText('3', { timeout: 10_000 })

    // the session now finds nothing under the old needle
    await expect(panel.locator('.fr-results tbody tr')).toHaveCount(0, { timeout: 10_000 })

    // Escape closes the panel
    await panel.locator('input').first().click()
    await sheets.keyboard.press('Escape')
    await expect(panel).toHaveCount(0, { timeout: 5_000 })
  } finally {
    await launched.app.close().catch(() => {})
  }
})
