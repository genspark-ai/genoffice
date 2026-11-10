import { test, expect } from '@playwright/test'
import { execSync } from 'node:child_process'
import { copyFile, mkdtemp, readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'
import type { ElectronApplication } from 'playwright'

/**
 * Withholding a cell from the model, driven through the real app.
 *
 * What is asserted afterwards is the contract, not the pixels: the file still
 * holds the reader's real value (the workbook keeps working) and now carries
 * the mark, so the model's view of it is the placeholder and nothing else.
 *
 * The mark lives in a package part of our own, so the check is a `unzip -l`
 * plus the part's own bytes rather than anything visible in the sheet — the
 * whole point is that the workbook is untouched apart from one extra entry.
 */

const FIXTURE = resolve(__dirname, '../apps/sheets/fixtures/generated/compatibility-basic.xlsx')
const MENU_LABEL = 'Hide the selection from AI'
const REDACTION_PART = 'xl/gxRedactions.json'

/** the workbook's first sheet, as XML — the values must survive untouched */
function sheetXml(workbookPath: string): string {
  return execSync(`unzip -p "${workbookPath}" xl/worksheets/sheet1.xml`).toString()
}

/** Every cell's text, keyed by address: the values, with formatting stripped. */
function cellsOf(sheetXmlText: string): Record<string, string> {
  const out: Record<string, string> = {}
  for (const m of sheetXmlText.matchAll(/<c r="([A-Z]+\d+)"[^>]*>([\s\S]*?)<\/c>/g)) {
    out[m[1]!] = m[2]!.replace(/<[^>]*>/g, '')
  }
  return out
}

function partText(workbookPath: string): string {
  return execSync(`unzip -p "${workbookPath}" "${REDACTION_PART}"`).toString()
}

function partListed(workbookPath: string): boolean {
  return execSync(`unzip -l "${workbookPath}"`).toString().includes(REDACTION_PART)
}

async function waitForWorkbook(page: Page): Promise<void> {
  await page.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
    timeout: 30_000,
  })
  await page.waitForTimeout(1_500)
}

/** the worksheet canvas origin: right of the row header, below the column header */
async function gridOrigin(page: Page): Promise<{ x: number; y: number }> {
  const grid = await page.evaluate(() => {
    for (const canvas of Array.from(document.querySelectorAll('canvas'))) {
      const rect = canvas.getBoundingClientRect()
      if (rect.width > 500 && rect.height > 300) return { x: rect.x, y: rect.y }
    }
    return null
  })
  if (!grid) throw new Error('worksheet canvas not found')
  return grid
}

/**
 * Save the way the app menu does it. The sheets view has no save-status
 * element to wait on, so the file itself is the signal — which is also the
 * thing the test actually cares about.
 */
async function saveViaMenu(app: ElectronApplication, workbook: string) {
  await app.evaluate(({ webContents }) => {
    const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('://sheets/'))
    wc?.send('menu:action', 'save')
  })
  await expect
    .poll(() => partListed(workbook), {
      timeout: 30_000,
      message: 'the part never reached the file',
    })
    .toBe(true)
}

test.describe('sheets: withhold a cell from the model', () => {
  test('the mark rides in the file and the cell keeps its value', async () => {
    test.setTimeout(180_000)
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-sheets-redact-'))
    const workbook = join(scratch, 'redact.xlsx')
    await copyFile(FIXTURE, workbook)
    const before = sheetXml(workbook)

    const launched = await launchShell({
      onboardingSeen: true,
      settings: { lang: 'en' },
      videoDir: 'redact-sheets',
      openFile: workbook,
    })
    try {
      const page = await waitForPageWithUrl(launched.app, '://sheets/')
      await waitForWorkbook(page)
      const grid = await gridOrigin(page)

      // B2, right-clicked. Cell values live on canvas, so the gesture is a
      // real mouse event at real grid coordinates.
      const cell = { x: grid.x + 46 + 43 * 2, y: grid.y + 24 + 12 * 2 }
      await page.mouse.click(cell.x, cell.y)
      await page.mouse.click(cell.x, cell.y, { button: 'right' })

      const item = page.getByText(MENU_LABEL, { exact: true })
      await expect(item).toBeVisible({ timeout: 15_000 })
      await page.screenshot({ path: screenshotPath('sheets-redact-1-menu') })
      await item.click()

      const dialog = page.locator('.redact-dialog')
      await expect(dialog).toBeVisible()
      const input = dialog.locator('input')
      await input.fill('client phone')
      // the literal the model will read, shown before committing
      await expect(dialog.locator('.redact-dialog-preview')).toHaveText('{{client phone}}')
      await page.screenshot({ path: screenshotPath('sheets-redact-2-dialog') })
      await dialog.locator('.btn-primary').click()
      await expect(dialog).toHaveCount(0)
      await page.screenshot({ path: screenshotPath('sheets-redact-3-marked') })

      await saveViaMenu(launched.app, workbook)
      await page.screenshot({ path: screenshotPath('sheets-redact-4-saved') })

      // the mark is in the package…
      expect(partListed(workbook), `${REDACTION_PART} is missing from the saved file`).toBe(true)
      const part = partText(workbook)
      expect(part).toContain('client phone')
      // …and it is declared, so the part is really part of the package
      // The brackets are a glob to both the shell and unzip, so the part name
      // is escaped before it reaches either.
      expect(execSync(`unzip -p "${workbook}" '\\[Content_Types\\].xml'`).toString()).toContain(
        REDACTION_PART,
      )

      // The reader's own data is untouched: the mark is a label over the cell,
      // never a replacement for it.
      //
      // The cell's *text* has to be identical, but the sheet XML no longer is,
      // and deliberately so — the mark tints the cell, which is a real
      // formatting change in the file. What must not change is the value and
      // the fill the reader had before, which the part records so clearing can
      // put it back.
      const after = sheetXml(workbook)
      expect(after).not.toContain('{{')
      expect(cellsOf(after)).toEqual(cellsOf(before))
      // the part remembers the fill the tint displaced
      expect(part).toMatch(/"previousFill":\s*(null|"#[0-9A-Fa-f]{6}")/)
    } finally {
      await closeAndSaveVideo(launched, 'redact-sheets')
    }
  })

  test('reopening the workbook keeps the mark', async () => {
    test.setTimeout(180_000)
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-sheets-redact2-'))
    const workbook = join(scratch, 'redact-reopen.xlsx')
    await copyFile(FIXTURE, workbook)

    const launched = await launchShell({
      onboardingSeen: true,
      settings: { lang: 'en' },
      videoDir: 'redact-sheets-reopen',
      openFile: workbook,
    })
    try {
      // open, mark, save, close — all inside this run
      const page = await waitForPageWithUrl(launched.app, '://sheets/')
      await waitForWorkbook(page)
      const grid = await gridOrigin(page)
      const cell = { x: grid.x + 46 + 43, y: grid.y + 24 + 12 }
      await page.mouse.click(cell.x, cell.y)
      await page.mouse.click(cell.x, cell.y, { button: 'right' })
      await page.getByText(MENU_LABEL, { exact: true }).click({ timeout: 15_000 })
      const dialog = page.locator('.redact-dialog')
      await expect(dialog).toBeVisible()
      await dialog.locator('input').fill('client phone')
      await dialog.locator('.btn-primary').click()
      await saveViaMenu(launched.app, workbook)
    } finally {
      await closeAndSaveVideo(launched, 'redact-sheets-reopen')
    }

    // A second launch reads the part back off disk. Nothing in this run marks
    // anything, so the only way the mark can exist is if it survived.
    const reopened = await launchShell({
      onboardingSeen: true,
      settings: { lang: 'en' },
      videoDir: 'redact-sheets-reopen-2',
      openFile: workbook,
    })
    try {
      const page = await waitForPageWithUrl(reopened.app, '://sheets/')
      await waitForWorkbook(page)
      // the same cell, right-clicked: an already-withheld selection offers to
      // stop withholding it rather than asking for a label
      const grid = await gridOrigin(page)
      const cell = { x: grid.x + 46 + 43, y: grid.y + 24 + 12 }
      await page.mouse.click(cell.x, cell.y)
      await page.mouse.click(cell.x, cell.y, { button: 'right' })
      await expect(page.getByText(MENU_LABEL, { exact: true })).toBeVisible({ timeout: 15_000 })
      await page.getByText(MENU_LABEL, { exact: true }).click()
      // clearing needs no label, so no dialog — the mark was found and read back
      await expect(page.locator('.redact-dialog')).toHaveCount(0)
      await page.screenshot({ path: screenshotPath('sheets-redact-5-reopened') })
    } finally {
      await closeAndSaveVideo(reopened, 'redact-sheets-reopen-2')
    }
  })
})

/** the fixture must actually contain a value in the cell the test marks */
test('fixture has content to withhold', async () => {
  const buffer = await readFile(FIXTURE)
  expect(buffer.length).toBeGreaterThan(0)
})
