import { test, expect } from '@playwright/test'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

process.env.GENOFFICE_DEBUG_HOOKS = '1'
const MOD = process.platform === 'darwin' ? 'Meta' : 'Control'

interface SheetFacade {
  getActiveWorkbook(): {
    getActiveSheet(): {
      getRange(
        row: number,
        column: number,
        rows: number,
        columns: number,
      ): {
        setValues(v: unknown[][]): Promise<unknown>
        activate(): unknown
        getValues(): unknown[][]
        getFormulas(): unknown[][]
      }
    }
  }
  executeCommand(id: string): Promise<unknown>
}

interface SheetDims {
  getActiveWorkbook(): { getActiveSheet(): { getMaxColumns(): number; getMaxRows(): number } }
}

test('paste special: transpose, add operation, paste link, undo', async () => {
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-paste-special-'))
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'sheets-paste-special' })
  try {
    const { app, page } = launched
    await app.evaluate(({ app: electronApp }, dir) => {
      electronApp.setPath('documents', dir)
    }, scratch)
    await expect(page.locator('.quick-card').nth(1)).toContainText('AI Sheets')
    await page.locator('.quick-card').nth(1).click()
    const sheets = await waitForPageWithUrl(app, '://sheets/')
    await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
      timeout: 30_000,
    })
    await sheets.waitForTimeout(1_500)
    const grid = await sheets.evaluate(() => {
      for (const canvas of document.querySelectorAll('canvas')) {
        const rect = canvas.getBoundingClientRect()
        if (rect.width > 500 && rect.height > 300) return { x: rect.x, y: rect.y }
      }
      return null
    })
    if (!grid) throw new Error('worksheet canvas not found')
    await sheets.mouse.click(grid.x + 46 + 43, grid.y + 24 + 12)

    const values = (r: number, c: number, rows: number, cols: number) =>
      sheets.evaluate(
        ([r, c, rows, cols]) => {
          const debug = (window as unknown as { __genofficeDebug: { univerAPI: SheetFacade } })
            .__genofficeDebug
          return debug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
            .getRange(r, c, rows, cols)
            .getValues()
        },
        [r, c, rows, cols] as const,
      )
    const activate = (r: number, c: number, rows: number, cols: number) =>
      sheets.evaluate(
        ([r, c, rows, cols]) => {
          const debug = (window as unknown as { __genofficeDebug: { univerAPI: SheetFacade } })
            .__genofficeDebug
          debug.univerAPI.getActiveWorkbook().getActiveSheet().getRange(r, c, rows, cols).activate()
        },
        [r, c, rows, cols] as const,
      )

    await sheets.evaluate(async () => {
      const debug = (window as unknown as { __genofficeDebug: { univerAPI: SheetFacade } })
        .__genofficeDebug
      const sheet = debug.univerAPI.getActiveWorkbook().getActiveSheet()
      await sheet.getRange(0, 0, 2, 2).setValues([
        [1, 2],
        [3, 4],
      ])
    })
    await activate(0, 0, 2, 2)
    await sheets.keyboard.press(`${MOD}+c`)
    await sheets.waitForTimeout(300)

    // transpose into D1
    await activate(0, 3, 1, 1)
    await sheets.keyboard.press('Control+Alt+v')
    const dialog = sheets.locator('.paste-special-dialog')
    await expect(dialog).toBeVisible()
    await sheets.screenshot({ path: 'test-results/paste-special-dialog.png' })
    await dialog.locator('.paste-special-flags input[type=checkbox]').nth(1).check()
    await dialog.locator('.dialog-actions .primary-action').click()
    await expect(dialog).toBeHidden()
    await expect(async () => {
      expect(await values(0, 3, 2, 2)).toEqual([
        [1, 3],
        [2, 4],
      ])
    }).toPass({ timeout: 10_000 })

    // add operation onto D1:E2
    await activate(0, 3, 1, 1)
    await sheets.keyboard.press('Control+Alt+v')
    await expect(dialog).toBeVisible()
    await dialog.locator('.paste-special-operations input[type=radio]').nth(1).check()
    await dialog.locator('.dialog-actions .primary-action').click()
    await expect(dialog).toBeHidden()
    await expect(async () => {
      expect(await values(0, 3, 2, 2)).toEqual([
        [2, 5],
        [5, 8],
      ])
    }).toPass({ timeout: 10_000 })

    // one undo step restores the transposed block
    await sheets.keyboard.press(`${MOD}+z`)
    await expect(async () => {
      expect(await values(0, 3, 2, 2)).toEqual([
        [1, 3],
        [2, 4],
      ])
    }).toPass({ timeout: 10_000 })

    // paste link into G1, opened through the context-menu command id
    await activate(0, 6, 1, 1)
    await sheets.evaluate(() => {
      const debug = (window as unknown as { __genofficeDebug: { univerAPI: SheetFacade } })
        .__genofficeDebug
      void debug.univerAPI.executeCommand('sheets.paste-special.open')
    })
    await expect(dialog).toBeVisible()
    await dialog.locator('.dialog-actions .secondary').first().click()
    await expect(dialog).toBeHidden()
    await expect(async () => {
      const formulas = await sheets.evaluate(() => {
        const debug = (window as unknown as { __genofficeDebug: { univerAPI: SheetFacade } })
          .__genofficeDebug
        return debug.univerAPI
          .getActiveWorkbook()
          .getActiveSheet()
          .getRange(0, 6, 2, 2)
          .getFormulas()
      })
      expect(formulas).toEqual([
        ['=A1', '=B1'],
        ['=A2', '=B2'],
      ])
    }).toPass({ timeout: 10_000 })
    await expect(async () => {
      expect(await values(0, 6, 2, 2)).toEqual([
        [1, 2],
        [3, 4],
      ])
    }).toPass({ timeout: 10_000 })
    await sheets.screenshot({ path: 'test-results/paste-special-after.png' })

    // operation against a formula target uses its computed value
    await sheets.evaluate(async () => {
      const debug = (window as unknown as { __genofficeDebug: { univerAPI: SheetFacade } })
        .__genofficeDebug
      await debug.univerAPI
        .getActiveWorkbook()
        .getActiveSheet()
        .getRange(3, 0, 1, 1)
        .setValues([['=A1*10']])
    })
    await expect(async () => {
      expect(await values(3, 0, 1, 1)).toEqual([[10]])
    }).toPass({ timeout: 10_000 })
    await activate(0, 0, 1, 1)
    await sheets.keyboard.press(`${MOD}+c`)
    await sheets.waitForTimeout(300)
    await activate(3, 0, 1, 1)
    await sheets.keyboard.press('Control+Alt+v')
    await expect(dialog).toBeVisible()
    await dialog.locator('.paste-special-operations input[type=radio]').nth(1).check()
    await dialog.locator('.dialog-actions .primary-action').click()
    await expect(dialog).toBeHidden()
    await expect(async () => {
      expect(await values(3, 0, 1, 1)).toEqual([[11]])
    }).toPass({ timeout: 10_000 })

    // a transposed block past the last column grows the sheet
    const maxColumns = await sheets.evaluate(() => {
      const debug = (window as unknown as { __genofficeDebug: { univerAPI: SheetDims } })
        .__genofficeDebug
      return debug.univerAPI.getActiveWorkbook().getActiveSheet().getMaxColumns()
    })
    await activate(0, 0, 2, 2)
    await sheets.keyboard.press(`${MOD}+c`)
    await sheets.waitForTimeout(300)
    await activate(5, maxColumns - 1, 1, 1)
    await sheets.keyboard.press('Control+Alt+v')
    await expect(dialog).toBeVisible()
    await dialog.locator('.paste-special-flags input[type=checkbox]').nth(1).check()
    await dialog.locator('.dialog-actions .primary-action').click()
    await expect(dialog).toBeHidden()
    await expect(async () => {
      expect(await values(5, maxColumns - 1, 2, 2)).toEqual([
        [1, 3],
        [2, 4],
      ])
    }).toPass({ timeout: 10_000 })
  } finally {
    await closeAndSaveVideo(launched, 'sheets-paste-special')
  }
})
