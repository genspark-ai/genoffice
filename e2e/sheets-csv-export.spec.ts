import { test, expect } from '@playwright/test'
import { existsSync } from 'node:fs'
import { copyFile, mkdtemp, readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

const FIXTURE = resolve(__dirname, '../apps/sheets/fixtures/generated/compatibility-basic.xlsx')

async function waitForWorkbook(page: Page): Promise<void> {
  await page.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
    timeout: 30_000,
  })
  await page.waitForTimeout(1_500)
}

test.describe('sheets: export the active sheet as CSV', () => {
  test('File > Export CSV writes a BOM-prefixed CSV of the grid', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-csv-e2e-'))
    const workbook = join(scratch, 'export-source.xlsx')
    const target = join(scratch, 'exported.csv')
    await copyFile(FIXTURE, workbook)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-csv-export',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, 'sheets/out')
      await waitForWorkbook(sheets)

      // Stub the native dialogs: pick the target path, answer "Continue as
      // CSV" (button index 1) if the formula-loss warning appears.
      await launched.app.evaluate(({ dialog }, csvPath) => {
        dialog.showSaveDialog = async () => ({ canceled: false, filePath: csvPath })
        dialog.showMessageBox = (async () => ({ response: 1, checkboxChecked: false })) as never
      }, target)

      // The export refuses politely until the workbook preload finishes, so
      // keep re-sending the menu action until the file lands.
      await expect(async () => {
        await launched.app.evaluate(({ webContents }) => {
          const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('sheets/out'))
          wc?.send('menu:action', 'export-csv')
        })
        expect(existsSync(target)).toBe(true)
      }).toPass({ timeout: 30_000, intervals: [2_000] })

      const bytes = await readFile(target)
      expect([bytes[0], bytes[1], bytes[2]]).toEqual([0xef, 0xbb, 0xbf])
      const text = bytes.subarray(3).toString('utf8')
      // the fixture's numeric cell renders as-is, rows end with CRLF
      expect(text).toContain('10')
      expect(text).toContain('\r\n')
    } finally {
      await closeAndSaveVideo(launched, 'sheets-csv-export')
    }
  })
})
