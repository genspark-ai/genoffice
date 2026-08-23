import { test, expect } from '@playwright/test'
import { execFileSync, execSync } from 'node:child_process'
import { copyFile, mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

const FIXTURE = resolve(__dirname, '../apps/sheets/fixtures/generated/compatibility-macro.xlsm')

const canReadClipboard = process.platform === 'darwin'

/** the leading backslashes stop unzip from reading [] as a match pattern */
function zipEntry(archive: string, name: string): Buffer {
  return execFileSync('unzip', ['-p', archive, name.replace(/([[\]])/g, '\\$1')])
}

async function waitForWorkbook(page: Page): Promise<void> {
  await page.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
    timeout: 30_000,
  })
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

test.describe('sheets: macro-enabled workbook (.xlsm)', () => {
  test('opens, edits, and saves with the VBA project preserved verbatim', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-xlsm-e2e-'))
    const workbook = join(scratch, 'macro.xlsm')
    await copyFile(FIXTURE, workbook)
    const vbaBefore = zipEntry(workbook, 'xl/vbaProject.bin')

    // ── session 1: open, edit A1, save in place (stays .xlsm) ──
    const first = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-xlsm',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(first.app, 'sheets/out')
      await waitForWorkbook(sheets)

      const a1 = await cellA1(sheets)
      await sheets.mouse.click(a1.x, a1.y)
      await expect(sheets.locator('.name-box')).toHaveValue('A1')

      await sheets.keyboard.type('MacroSafe', { delay: 50 })
      await sheets.keyboard.press('Enter')
      await sheets.screenshot({ path: screenshotPath('sheets-xlsm-edited') })

      await first.app.evaluate(({ webContents }) => {
        const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('sheets/out'))
        wc?.send('menu:action', 'save')
      })
      await expect(() => {
        expect(zipEntry(workbook, 'xl/worksheets/sheet1.xml').toString()).toContain('MacroSafe')
      }).toPass({ timeout: 15_000 })

      // macros are never executed, but they must survive the save byte-for-byte
      expect(zipEntry(workbook, 'xl/vbaProject.bin').equals(vbaBefore)).toBe(true)
      const contentTypes = zipEntry(workbook, '[Content_Types].xml').toString()
      expect(contentTypes).toContain('application/vnd.ms-excel.sheet.macroEnabled.main+xml')
      expect(contentTypes).toContain('application/vnd.ms-office.vbaProject')
      const rels = zipEntry(workbook, 'xl/_rels/workbook.xml.rels').toString()
      expect(rels).toContain('vbaProject.bin')
    } finally {
      await closeAndSaveVideo(first, 'sheets-xlsm')
    }

    // ── session 2: the saved .xlsm reopens with the edit in the grid ──
    const second = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-xlsm-reopen',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(second.app, 'sheets/out')
      await waitForWorkbook(sheets)

      const a1 = await cellA1(sheets)
      await sheets.mouse.click(a1.x, a1.y)
      await expect(sheets.locator('.name-box')).toHaveValue('A1')
      if (canReadClipboard) {
        expect(await copyActiveCell(sheets)).toBe('MacroSafe')
      }
      await sheets.screenshot({ path: screenshotPath('sheets-xlsm-reopened') })
    } finally {
      await closeAndSaveVideo(second, 'sheets-xlsm-reopen')
    }
  })
})
