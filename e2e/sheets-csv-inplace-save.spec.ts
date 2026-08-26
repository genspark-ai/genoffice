import { test, expect } from '@playwright/test'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { readdirSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

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

test.describe('sheets: a CSV keeps its identity through Save', () => {
  test('Cmd+S writes the edit back to the original .csv', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-csv-save-e2e-'))
    const csvSource = join(scratch, 'data.csv')
    await writeFile(csvSource, 'name,amount\r\nalpha,10\r\n')

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-csv-inplace-save',
      openFile: csvSource,
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, 'sheets/out')
      await waitForWorkbook(sheets)

      const a1 = await cellA1(sheets)
      await sheets.mouse.click(a1.x, a1.y)
      await expect(sheets.locator('.name-box')).toHaveValue('A1')
      await sheets.keyboard.type('Hello', { delay: 50 })
      await sheets.keyboard.press('Enter')

      // Answer the "keep this format?" question with the default: Continue as CSV.
      await launched.app.evaluate(({ dialog }) => {
        dialog.showMessageBox = (async () => ({
          response: 0,
          checkboxChecked: false,
        })) as never
      })

      // Save may be refused politely until the preload finishes — retry.
      await expect(async () => {
        await launched.app.evaluate(({ webContents }) => {
          const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('sheets/out'))
          wc?.send('menu:action', 'save')
        })
        const bytes = await readFile(csvSource)
        expect(bytes.subarray(3).toString('utf8')).toContain('Hello')
      }).toPass({ timeout: 30_000, intervals: [2_000] })

      const bytes = await readFile(csvSource)
      expect([bytes[0], bytes[1], bytes[2]]).toEqual([0xef, 0xbb, 0xbf])
      const text = bytes.subarray(3).toString('utf8')
      expect(text).toContain('Hello,amount')
      expect(text).toContain('alpha,10')
      // No Save As detour: the folder holds only the .csv, no stray .xlsx.
      expect(readdirSync(scratch)).toEqual(['data.csv'])
    } finally {
      await closeAndSaveVideo(launched, 'sheets-csv-inplace-save')
    }
  })
})
