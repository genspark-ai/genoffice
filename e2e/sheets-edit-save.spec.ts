import { test, expect } from '@playwright/test'
import { execSync } from 'node:child_process'
import { copyFile, mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

const FIXTURE = resolve(__dirname, '../apps/sheets/fixtures/generated/compatibility-basic.xlsx')

/** cell values live on canvas, so reading them back goes through the system clipboard */
const canReadClipboard = process.platform === 'darwin'

async function waitForWorkbook(page: Page): Promise<void> {
  // the sheet tab strip is DOM; the fixture's only sheet is "Sheet1"
  await page.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
    timeout: 30_000,
  })
  // let the first viewport range stream in before dispatching mouse events
  await page.waitForTimeout(1_500)
}

/** center of cell A1: right of the ~46px row header, below the ~24px column header */
async function cellA1(page: Page): Promise<{ x: number; y: number }> {
  const grid = await page.evaluate(() => {
    for (const canvas of document.querySelectorAll('canvas')) {
      const rect = canvas.getBoundingClientRect()
      if (rect.width > 500 && rect.height > 300) return { x: rect.x, y: rect.y }
    }
    return null
  })
  if (!grid) throw new Error('worksheet canvas not found')
  return { x: grid.x + 46 + 43, y: grid.y + 24 + 12 }
}

async function copyActiveCell(page: Page): Promise<string> {
  await page.keyboard.press('Meta+c')
  await page.waitForTimeout(500)
  return execSync('pbpaste').toString()
}

function sheetXml(workbookPath: string): string {
  return execSync(`unzip -p "${workbookPath}" xl/worksheets/sheet1.xml`).toString()
}

test.describe('sheets: edit and save an external workbook', () => {
  test('cell edit round-trips through save and reopen', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-sheets-e2e-'))
    const workbook = join(scratch, 'edit-save.xlsx')
    await copyFile(FIXTURE, workbook)

    // ── session 1: open, edit A1, save ──
    const first = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-edit-save',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(first.app, 'sheets/out')
      await waitForWorkbook(sheets)

      const a1 = await cellA1(sheets)
      await sheets.mouse.click(a1.x, a1.y)
      await expect(sheets.locator('.name-box')).toHaveValue('A1')

      await sheets.keyboard.type('Hello', { delay: 50 })
      await sheets.keyboard.press('Enter')

      if (canReadClipboard) {
        await sheets.mouse.click(a1.x, a1.y)
        expect(await copyActiveCell(sheets)).toBe('Hello')
      }
      await sheets.screenshot({ path: screenshotPath('sheets-edited') })

      // File > Save, routed to the sheets view the same way the app menu does it
      await first.app.evaluate(({ webContents }) => {
        const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('sheets/out'))
        wc?.send('menu:action', 'save')
      })
      await expect(() => {
        expect(sheetXml(workbook)).toContain('Hello')
      }).toPass({ timeout: 15_000 })
      // surgical save: the untouched sibling cell survives byte-identical
      expect(sheetXml(workbook)).toContain('<v>10</v>')
    } finally {
      await closeAndSaveVideo(first, 'sheets-edit-save')
    }

    // ── session 2: reopen the saved file, the edited value is really in the grid ──
    const second = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-reopen',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(second.app, 'sheets/out')
      await waitForWorkbook(sheets)

      const a1 = await cellA1(sheets)
      await sheets.mouse.click(a1.x, a1.y)
      await expect(sheets.locator('.name-box')).toHaveValue('A1')
      if (canReadClipboard) {
        expect(await copyActiveCell(sheets)).toBe('Hello')
      }
      await sheets.screenshot({ path: screenshotPath('sheets-reopened') })
    } finally {
      await closeAndSaveVideo(second, 'sheets-reopen')
    }
  })
})
