import { test, expect } from '@playwright/test'
import { existsSync } from 'node:fs'
import { copyFile, mkdtemp, readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

const FIXTURE = resolve(__dirname, '../apps/sheets/fixtures/generated/compatibility-basic.xlsx')
// has a cached formula (B3 =SUM(C1:C2)) so the CSV flows hit the loss warning;
// its single sheet is named "Data"
const FORMULA_FIXTURE = resolve(
  __dirname,
  '../apps/sheets/fixtures/generated/compatibility-edit.xlsx',
)

async function waitForWorkbook(page: Page, sheetName = 'Sheet1'): Promise<void> {
  await page.waitForFunction((name) => document.body.textContent?.includes(name), sheetName, {
    timeout: 30_000,
  })
  await page.waitForTimeout(1_500)
}

/// Stubs the native dialogs: the save picker returns `csvPath`, and every
/// message box answers with its "Continue as CSV" button, recording each box
/// in `globalThis.__boxes` so tests can assert which warnings actually fired.
async function stubDialogs(
  app: import('@playwright/test').ElectronApplication,
  csvPath: string,
): Promise<void> {
  await app.evaluate(({ dialog }, target) => {
    const g = globalThis as unknown as { __boxes: { message: string }[] }
    g.__boxes = []
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: target })
    dialog.showMessageBox = (async (...args: unknown[]) => {
      const options = (
        args.length === 1 ? args[0] : args[1]
      ) as import('electron').MessageBoxOptions
      g.__boxes.push({ message: options.message ?? '' })
      const idx = (options.buttons ?? []).indexOf('Continue as CSV')
      return { response: idx >= 0 ? idx : 1, checkboxChecked: false }
    }) as never
  }, csvPath)
}

async function shownBoxes(
  app: import('@playwright/test').ElectronApplication,
): Promise<{ message: string }[]> {
  return app.evaluate(() => (globalThis as unknown as { __boxes: { message: string }[] }).__boxes)
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

  test('Export CSV with formulas: Continue on the loss warning writes the file and toasts', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-csv-formula-e2e-'))
    const workbook = join(scratch, 'formula-source.xlsx')
    const target = join(scratch, 'exported.csv')
    await copyFile(FORMULA_FIXTURE, workbook)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-csv-export-formula',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, 'sheets/out')
      await waitForWorkbook(sheets, 'Data')
      await stubDialogs(launched.app, target)

      await expect(async () => {
        await launched.app.evaluate(({ webContents }) => {
          const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('sheets/out'))
          wc?.send('menu:action', 'export-csv')
        })
        expect(existsSync(target)).toBe(true)
      }).toPass({ timeout: 30_000, intervals: [2_000] })

      // the warning must actually have fired — a formula-free fixture would
      // silently skip the branch this test exists for
      const boxes = await shownBoxes(launched.app)
      expect(boxes.map((box) => box.message).join('\n')).toContain('formulas')

      // formula cell exported as its computed value
      const text = (await readFile(target)).subarray(3).toString('utf8')
      expect(text).toContain('5')

      // the export announces itself with a toast, not just the status bar
      await launched.app.evaluate(({ webContents }) => {
        const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('sheets/out'))
        wc?.send('menu:action', 'export-csv')
      })
      await sheets.waitForSelector('.app-toast', { timeout: 10_000 })
      expect(await sheets.locator('.app-toast').textContent()).toContain('Exported')
    } finally {
      await closeAndSaveVideo(launched, 'sheets-csv-export-formula')
    }
  })

  test('Save As with a .csv pick rides the formula warning through to the file', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-csv-saveas-e2e-'))
    const workbook = join(scratch, 'saveas-source.xlsx')
    const target = join(scratch, 'saved.csv')
    await copyFile(FORMULA_FIXTURE, workbook)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-csv-saveas',
      openFile: workbook,
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, 'sheets/out')
      await waitForWorkbook(sheets, 'Data')
      await stubDialogs(launched.app, target)

      await expect(async () => {
        await launched.app.evaluate(({ webContents }) => {
          const wc = webContents.getAllWebContents().find((w) => w.getURL().includes('sheets/out'))
          wc?.send('menu:action', 'save-as')
        })
        expect(existsSync(target)).toBe(true)
      }).toPass({ timeout: 30_000, intervals: [2_000] })

      const boxes = await shownBoxes(launched.app)
      expect(boxes.map((box) => box.message).join('\n')).toContain('formulas')

      const bytes = await readFile(target)
      expect([bytes[0], bytes[1], bytes[2]]).toEqual([0xef, 0xbb, 0xbf])
      expect(bytes.subarray(3).toString('utf8')).toContain('5')
    } finally {
      await closeAndSaveVideo(launched, 'sheets-csv-saveas')
    }
  })
})
